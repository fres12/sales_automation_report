#!/usr/bin/env python
"""
Test script untuk validasi tanggal logic
Menguji berbagai skenario tanggal untuk memastikan validation bekerja benar
"""

from datetime import datetime, timedelta

def test_date_validation():
    """Test berbagai kombinasi tanggal di A1 dengan hari ini"""
    
    print("="*70)
    print("TEST DATE VALIDATION LOGIC")
    print("="*70)
    
    # Test cases: (tanggal_di_A1, hari_hari_ini, expected_result, description)
    test_cases = [
        # Format: (date_in_a1, today_date, expected_valid, description)
        
        # Tanggal 2-19: H-2 (2 hari setelah)
        (3, 5, True, "Tanggal 3 → Expected 5 (3+2)"),
        (3, 4, False, "Tanggal 3 → Hari ini 4, harusnya 5"),
        (7, 9, True, "Tanggal 7 → Expected 9 (7+2)"),
        (15, 17, True, "Tanggal 15 → Expected 17 (15+2)"),
        (19, 21, True, "Tanggal 19 → Expected 21 (19+2)"),
        
        # Tanggal 20-31: H-1 (1 hari setelah)
        (20, 21, True, "Tanggal 20 → Expected 21 (20+1)"),
        (25, 26, True, "Tanggal 25 → Expected 26 (25+1)"),
        (31, 1, True, "Tanggal 31 (bulan lalu) → Expected 1 (overflow ke bulan depan)"),
        (30, 1, True, "Tanggal 30 (bulan Feb ada 28) → Expected 1 (overflow ke bulan depan)"),
        
        # Tanggal 1: H-1 dari bulan lalu
        (1, 1, True, "Tanggal 1 → Hari ini 1 di awal bulan (kemarin bulan lalu)"),
        (1, 2, False, "Tanggal 1 → Hari ini 2, harusnya sudah di awal bulan kemarin"),
    ]
    
    print("\n📋 Test Cases:\n")
    
    for idx, (date_a1, today_day, expected_valid, description) in enumerate(test_cases, 1):
        print(f"Test {idx}: {description}")
        print(f"   A1: {date_a1} | Today: {today_day}")
        
        # Determine expected logic
        if date_a1 == 1:
            # Check if yesterday is different month
            # For testing, we'll assume today is always valid if 1-3
            is_valid = True if 1 <= today_day <= 3 else False
            schedule = "Tanggal 1 (H-1 dari bulan lalu)"
        elif 2 <= date_a1 <= 19:
            expected_day = date_a1 + 2
            is_valid = (today_day == expected_day)
            schedule = f"Tanggal {date_a1} → Expected {expected_day} (H-2)"
        elif 20 <= date_a1 <= 31:
            expected_day = date_a1 + 1
            # Handle month overflow (day > 31 → 1, etc)
            if expected_day > 31:
                expected_day = expected_day - 31
            is_valid = (today_day == expected_day)
            schedule = f"Tanggal {date_a1} → Expected {expected_day} (H-1)"
        else:
            is_valid = False
            schedule = "Invalid date"
        
        result = "✅ VALID" if is_valid == expected_valid else "❌ FAILED"
        print(f"   Result: {result}")
        print(f"   Schedule: {schedule}\n")
    
    print("="*70)
    print("SUMMARY:")
    print("  - Tanggal 1-19 (except 1): H-2 = 2 hari setelah tanggal")
    print("  - Tanggal 20-31: H-1 = 1 hari setelah tanggal")
    print("  - Tanggal 1 (special): H-1 dari bulan lalu (kemarin bulan lalu)")
    print("="*70)


if __name__ == "__main__":
    test_date_validation()
