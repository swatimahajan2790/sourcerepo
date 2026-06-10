# -*- coding: utf-8 -*-
"""
Quick test to verify subprocess execution in the external environment.
Run this from IronPython to check if CPython subprocess calls work.
"""

import os
import sys

# Add modules to path
_modules_dir = os.path.dirname(os.path.abspath(__file__))
if _modules_dir not in sys.path:
    sys.path.insert(0, _modules_dir)

from _imports import _CPYTHON, _PDF_HELPER, _LIB_DIR, _subprocess, _find_fallback_cpython
from logger import LogManager


def test_python_executable():
    """Test 1: Can we even call the Python executable?"""
    print("\n" + "="*70)
    print("TEST 1: Python Executable Accessibility")
    print("="*70)
    
    if not _CPYTHON:
        print("✗ FAIL: _CPYTHON is None - no Python found")
        return False
    
    print("  Python path: {}".format(_CPYTHON))
    print("  File exists: {}".format(os.path.isfile(_CPYTHON)))
    
    if not os.path.isfile(_CPYTHON):
        print("✗ FAIL: Python executable not found at path")
        return False
    
    # Try calling --version
    try:
        _creationflags = 0x08000000 if os.name == 'nt' else 0
        proc = _subprocess.Popen(
            [_CPYTHON, "--version"],
            stdout=_subprocess.PIPE,
            stderr=_subprocess.PIPE,
            creationflags=_creationflags,
            shell=False
        )
        out, err = proc.communicate()
        
        version_str = (out.decode("utf-8", "replace") + err.decode("utf-8", "replace")).strip()
        print("  Version output: {}".format(version_str))
        
        if proc.returncode == 0:
            print("✓ PASS: Python executable is callable")
            return True
        else:
            print("✗ FAIL: Python returned code {}".format(proc.returncode))
            return False
    except Exception as ex:
        print("✗ FAIL: Exception calling Python: {}".format(ex))
        return False


def test_pymupdf():
    """Test 2: Is PyMuPDF (fitz) installed in the Python?"""
    print("\n" + "="*70)
    print("TEST 2: PyMuPDF Availability")
    print("="*70)
    
    try:
        _creationflags = 0x08000000 if os.name == 'nt' else 0
        proc = _subprocess.Popen(
            [_CPYTHON, "-c", "import fitz; print('PyMuPDF {} OK'.format(fitz.__version__))"],
            stdout=_subprocess.PIPE,
            stderr=_subprocess.PIPE,
            creationflags=_creationflags,
            shell=False
        )
        out, err = proc.communicate()
        
        if proc.returncode == 0:
            print("  {}".format(out.decode("utf-8", "replace").strip()))
            print("✓ PASS: PyMuPDF is available")
            return True
        else:
            err_msg = err.decode("utf-8", "replace").strip()
            print("  Error: {}".format(err_msg))
            print("✗ FAIL: PyMuPDF not available (return code {})".format(proc.returncode))
            return False
    except Exception as ex:
        print("✗ FAIL: Exception testing PyMuPDF: {}".format(ex))
        return False


def test_pdf_helper():
    """Test 3: Can we call the PDF helper script?"""
    print("\n" + "="*70)
    print("TEST 3: PDF Helper Script Call")
    print("="*70)
    
    print("  Helper path: {}".format(_PDF_HELPER))
    print("  File exists: {}".format(os.path.isfile(_PDF_HELPER)))
    
    if not os.path.isfile(_PDF_HELPER):
        print("✗ FAIL: Helper script not found")
        return False
    
    print("  LIB_DIR: {}".format(_LIB_DIR))
    
    # Try calling help/usage
    try:
        _creationflags = 0x08000000 if os.name == 'nt' else 0
        proc = _subprocess.Popen(
            [_CPYTHON, _PDF_HELPER],  # Call without args to get usage
            stdout=_subprocess.PIPE,
            stderr=_subprocess.PIPE,
            creationflags=_creationflags,
            shell=False
        )
        out, err = proc.communicate()
        
        # Should fail with usage message
        usage = err.decode("utf-8", "replace").strip()
        if "Usage:" in usage or "convert" in usage:
            print("  {}".format(usage))
            print("✓ PASS: Helper script is callable and responsive")
            return True
        else:
            print("  Output: {}".format(out.decode("utf-8", "replace")[:200]))
            print("  Error: {}".format(usage[:200]))
            print("✗ FAIL: Helper script did not return expected usage message")
            return False
    except Exception as ex:
        print("✗ FAIL: Exception calling helper: {}".format(ex))
        return False


def test_fallback_python():
    """Test 4: Is there a system Python available as fallback?"""
    print("\n" + "="*70)
    print("TEST 4: Fallback Python Search")
    print("="*70)
    
    try:
        fallback = _find_fallback_cpython()
        if fallback:
            print("  Found: {}".format(fallback))
            print("  File exists: {}".format(os.path.isfile(fallback)))
            
            # Verify it has PyMuPDF
            _creationflags = 0x08000000 if os.name == 'nt' else 0
            proc = _subprocess.Popen(
                [fallback, "-c", "import fitz; print('OK')"],
                stdout=_subprocess.PIPE,
                stderr=_subprocess.PIPE,
                creationflags=_creationflags,
                shell=False
            )
            out, err = proc.communicate()
            
            if proc.returncode == 0:
                print("✓ PASS: Fallback Python exists and has PyMuPDF")
                return True
            else:
                print("⚠ WARNING: Fallback Python found but may lack PyMuPDF")
                return False
        else:
            print("⚠ WARNING: No system Python with PyMuPDF found (fallback unavailable)")
            return False
    except Exception as ex:
        print("✗ FAIL: Exception testing fallback: {}".format(ex))
        return False


def main():
    """Run all tests."""
    print("\n\n")
    print("#"*70)
    print("# DOCLINK SUBPROCESS DIAGNOSTIC TEST")
    print("#"*70)
    
    results = {
        "Python Executable": test_python_executable(),
        "PyMuPDF": test_pymupdf(),
        "PDF Helper": test_pdf_helper(),
        "Fallback Python": test_fallback_python(),
    }
    
    print("\n\n" + "="*70)
    print("SUMMARY")
    print("="*70)
    
    for test_name, passed in results.items():
        status = "✓ PASS" if passed else "✗ FAIL"
        print("  {}: {}".format(test_name, status))
    
    all_pass = all(results.values())
    print("\nOverall: {}".format("✓ ALL TESTS PASSED" if all_pass else "✗ SOME TESTS FAILED"))
    
    return all_pass


if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
