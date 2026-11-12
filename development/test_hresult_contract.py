"""HRESULT/VARIANT ABI models and source regression guards, not VBA execution.

The original native contract incorrectly passed negative MAPI HRESULTs to
CVErr, whose numeric input range is 0..65535. The byte models use the documented
Windows VARIANT header/union sizes for x86 and x64; they do not claim to execute
Outlook or the VBA compiler. Native startup tests also exercise the real helper.

Primary contracts:
https://learn.microsoft.com/en-us/openspecs/microsoft_general_purpose_programming_languages/ms-vbal/7a7facef-2f52-4d08-918e-63bf6a9261ed
https://learn.microsoft.com/en-us/windows/win32/api/oaidl/ns-oaidl-variant
"""

from __future__ import annotations

import ctypes
from pathlib import Path
import re
import struct
import unittest


ROOT = Path(__file__).resolve().parent
LEGACY = (ROOT / "mb_legacy.basfrag").read_text(encoding="ascii")
DECLARATIONS = (ROOT / "mb_legacy_declarations.basfrag").read_text(encoding="ascii")
FIDELITY = (ROOT / "mb_fidelity.basfrag").read_text(encoding="ascii")
NOT_FOUND = -2147221233  # 0x8004010F
ACCESS_DENIED = -2147024891  # 0x80070005
VT_ERROR = 10


def function(name: str) -> str:
    match = re.search(r"(?ms)^Private Function " + re.escape(name)
                      + r"\b.*?^End Function$", LEGACY)
    if match is None:
        raise AssertionError(f"Missing private function: {name}")
    return match.group()


def vba_cverr(code: int) -> int:
    """Model the MS-VBAL numeric-input precondition, including old failure."""
    if not 0 <= code <= 65535:
        raise ValueError("VBA error 5: invalid procedure call or argument")
    return code


def variant_type(bits: int):
    """Windows VARIANT union includes an 8-byte scalar and two record pointers."""
    pointer = ctypes.c_uint32 if bits == 32 else ctypes.c_uint64

    class Record(ctypes.Structure):
        _fields_ = [("record", pointer), ("info", pointer)]

    class Value(ctypes.Union):
        _fields_ = [("scalar", ctypes.c_int64), ("scode", ctypes.c_int32),
                    ("record", Record)]

    class Variant(ctypes.Structure):
        _fields_ = [("vt", ctypes.c_uint16), ("reserved1", ctypes.c_uint16),
                    ("reserved2", ctypes.c_uint16), ("reserved3", ctypes.c_uint16),
                    ("value", Value)]

    return Variant


def owned_variant(bits: int, code: int, vartype: int = VT_ERROR) -> bytes:
    data = bytearray(ctypes.sizeof(variant_type(bits)))
    struct.pack_into("<H", data, 0, vartype)
    struct.pack_into("<i", data, 8, code)
    return bytes(data)


def exact_scode(data: bytes) -> int:
    if struct.unpack_from("<H", data)[0] != VT_ERROR:
        raise ValueError("Only an owned scalar VT_ERROR is readable")
    return struct.unpack_from("<i", data, 8)[0]


class NumericBoundaryCases(unittest.TestCase):
    def test_original_not_found_numeric_cverr_call_is_invalid(self):
        with self.assertRaisesRegex(ValueError, "error 5"):
            vba_cverr(NOT_FOUND)

    def test_original_access_denied_numeric_cverr_call_is_invalid(self):
        with self.assertRaisesRegex(ValueError, "error 5"):
            vba_cverr(ACCESS_DENIED)

    def test_cverr_accepts_zero_and_65535_only_within_its_domain(self):
        self.assertEqual(vba_cverr(0), 0)
        self.assertEqual(vba_cverr(65535), 65535)
        for value in (-1, 65536, 2147483647):
            with self.subTest(value=value), self.assertRaises(ValueError):
                vba_cverr(value)

    def test_not_found_and_access_denied_remain_distinct(self):
        for bits in (32, 64):
            self.assertEqual(exact_scode(owned_variant(bits, NOT_FOUND)), NOT_FOUND)
            self.assertEqual(exact_scode(owned_variant(bits, ACCESS_DENIED)), ACCESS_DENIED)
            self.assertNotEqual(NOT_FOUND, ACCESS_DENIED)

    def test_matching_low16_bits_never_establish_missing_property(self):
        low16 = NOT_FOUND & 0xFFFF
        self.assertEqual(low16, 271)
        self.assertNotEqual(exact_scode(owned_variant(64, low16)), NOT_FOUND)

    def test_different_facility_with_same_low16_is_not_not_found(self):
        different_facility = -2147024625  # 0x8007010F
        self.assertEqual(different_facility & 0xFFFF, NOT_FOUND & 0xFFFF)
        self.assertNotEqual(exact_scode(owned_variant(32, different_facility)), NOT_FOUND)


class VariantAbiCases(unittest.TestCase):
    def test_x86_header_offset_and_full_structure_size(self):
        value = variant_type(32)
        self.assertEqual(value.value.offset, 8)
        self.assertEqual(ctypes.sizeof(value), 16)

    def test_x64_header_offset_and_full_structure_size(self):
        value = variant_type(64)
        self.assertEqual(value.value.offset, 8)
        self.assertEqual(ctypes.sizeof(value), 24)

    def test_scode_stays_four_bytes_on_both_architectures(self):
        self.assertEqual(ctypes.sizeof(ctypes.c_int32), 4)
        for bits in (32, 64):
            for code in (0, 65535, 2147483647, -2147483648, -1, NOT_FOUND, ACCESS_DENIED):
                with self.subTest(bits=bits, code=code):
                    self.assertEqual(exact_scode(owned_variant(bits, code)), code)

    def test_scode_patch_preserves_entire_other_x86_storage(self):
        before = bytearray(b"\xA5" * ctypes.sizeof(variant_type(32)))
        struct.pack_into("<H", before, 0, VT_ERROR)
        after = before.copy()
        struct.pack_into("<i", after, 8, NOT_FOUND)
        self.assertEqual(after[:8] + after[12:], before[:8] + before[12:])

    def test_scode_patch_preserves_entire_other_x64_storage(self):
        before = bytearray(b"\xA5" * ctypes.sizeof(variant_type(64)))
        struct.pack_into("<H", before, 0, VT_ERROR)
        after = before.copy()
        struct.pack_into("<i", after, 8, ACCESS_DENIED)
        self.assertEqual(after[:8] + after[12:], before[:8] + before[12:])

    def test_raw_reader_rejects_nonerror_variant_tags(self):
        for tag in (0, 1, 3, 8, 9, 11, 12, 17):
            with self.subTest(tag=tag), self.assertRaises(ValueError):
                exact_scode(owned_variant(64, NOT_FOUND, tag))

    def test_raw_reader_rejects_byref_and_array_error_tags(self):
        for tag in (VT_ERROR | 0x4000, VT_ERROR | 0x2000):
            with self.subTest(tag=tag), self.assertRaises(ValueError):
                exact_scode(owned_variant(32, NOT_FOUND, tag))


class SourceRegressionCases(unittest.TestCase):
    def test_constructor_initializes_error_before_only_four_byte_patch(self):
        body = function("MBL_HResultError")
        self.assertIn("Private Function MBL_HResultError(ByVal code As Long) As Variant", body)
        initialize = body.index("value = CVErr(0)")
        patch = body.index("MBL_CopyErrorMemory VarPtr(value) + 8, VarPtr(code), 4")
        self.assertLess(initialize, patch)
        self.assertIn("If MBL_ErrorHResult(value) <> code Then Err.Raise", body)
        self.assertEqual(len(re.findall(r"(?m)^\s*MBL_CopyErrorMemory\b", body)), 1)

    def test_reader_copies_and_checks_scalar_tag_before_reading_scode(self):
        body = function("MBL_ErrorHResult")
        self.assertIn("Private Function MBL_ErrorHResult(ByVal value As Variant) As Long", body)
        logical_check = body.index("If VarType(value) <> vbError Then Err.Raise")
        local_copy = body.index("localValue = value")
        stored_check = body.index("If storedType <> vbError Then Err.Raise")
        read = body.index("MBL_CopyErrorMemory VarPtr(result), VarPtr(localValue) + 8, 4")
        self.assertLess(logical_check, local_copy)
        self.assertLess(local_copy, stored_check)
        self.assertLess(stored_check, read)
        self.assertIn("MBL_CopyErrorMemory VarPtr(storedType), VarPtr(localValue), 2", body)

    def test_memory_declaration_is_private_and_pointer_width_safe(self):
        self.assertRegex(DECLARATIONS, r"(?s)#If VBA7 Then\s+Private Declare PtrSafe Sub MBL_CopyErrorMemory.*?"
                         r"ByVal destination As LongPtr, ByVal source As LongPtr, ByVal bytes As LongPtr")
        self.assertRegex(DECLARATIONS, r"(?s)#Else\s+Private Declare Sub MBL_CopyErrorMemory.*?"
                         r"ByVal destination As Long, ByVal source As Long, ByVal bytes As Long")
        self.assertEqual(DECLARATIONS.count('Alias "RtlMoveMemory"'), 2)

    def test_all_numeric_cverr_calls_have_literal_inputs_in_valid_range(self):
        calls = []
        for source in ROOT.glob("*.basfrag"):
            code = "\n".join(line.split("'", 1)[0] for line in source.read_text(encoding="ascii").splitlines())
            calls.extend((source.name, value.strip()) for value in re.findall(r"\bCVErr\(([^()]*)\)", code))
        self.assertTrue(calls)
        for filename, value in calls:
            with self.subTest(filename=filename, value=value):
                self.assertRegex(value, r"^\d+$")
                self.assertGreaterEqual(int(value), 0)
                self.assertLessEqual(int(value), 65535)

    def test_production_missing_property_compares_exact_hresult(self):
        body = function("MBL_MissingProperty")
        self.assertIn("If Not IsError(value) Then Exit Function", body)
        self.assertIn("code = MBL_ErrorHResult(value)", body)
        self.assertIn("If code = MBL_MAPI_NOT_FOUND Then", body)
        self.assertIn("Err.Raise MB_ERR + 178", body)
        self.assertNotIn("CVErr(", body)
        self.assertNotRegex(body, r"(?i)\b(?:And\s+&HFFFF|CStr\(value\)\s*=)")

    def test_native_property_contract_covers_signed_boundaries_and_low16_alias(self):
        body = function("MBL_PropertySelfTest")
        self.assertIn("absent = MBL_HResultError(MBL_MAPI_NOT_FOUND)", body)
        self.assertIn("denied = MBL_HResultError(&H80070005)", body)
        self.assertIn("MBL_ErrorHResult(absent) <> MBL_MAPI_NOT_FOUND", body)
        self.assertIn("MBL_ErrorHResult(denied) <> &H80070005", body)
        self.assertIn("Array(0&, 65535&, &H7FFFFFFF, &H80000000, &HFFFFFFFF)", body)
        self.assertIn('MBL_TestRejectProperty(CVErr(271), "ansi")', body)
        self.assertIn('MBL_TestRejectProperty(MBL_HResultError(&H8007010F), "ansi")', body)

    def test_category_test_uses_same_full_hresult_constructor(self):
        self.assertIn("missing = MBL_HResultError(MBL_MAPI_NOT_FOUND)", FIDELITY)
        self.assertIn('MBE_CategoryRejects(MBL_HResultError(&H80070005), "")', FIDELITY)
        self.assertNotIn("CVErr(", FIDELITY)

    def test_memory_access_is_limited_to_owned_local_storage(self):
        lines = [line.strip() for line in LEGACY.splitlines()
                 if line.strip().startswith("MBL_CopyErrorMemory ")]
        self.assertEqual(lines, [
            "MBL_CopyErrorMemory VarPtr(storedType), VarPtr(localValue), 2",
            "MBL_CopyErrorMemory VarPtr(result), VarPtr(localValue) + 8, 4",
            "MBL_CopyErrorMemory VarPtr(value) + 8, VarPtr(code), 4",
        ])
        for name in ("MBL_ErrorHResult", "MBL_HResultError"):
            body = function(name)
            self.assertNotRegex(body, r"(?i)\.\s*(?:Save|Send|Delete|Move|SaveAs)\b")


if __name__ == "__main__":
    unittest.main(verbosity=2)
