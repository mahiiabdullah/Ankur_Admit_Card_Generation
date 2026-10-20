# Exam Roll Number Generator — Instructions for Claude Code

This project generates exam roll numbers for students.
Read this whole file before generating any roll.

---

## 1. Roll format

Every roll is **exactly 8 digits**:

```
[C][S][G][2][6][N][N][N]
 │  │  │  └─┬─┘ └──┬──┘
 │  │  │    │      └── Serial: 3 digits, zero-padded (001–999)
 │  │  │    └───────── Year: always "26"
 │  │  └────────────── Gender code
 │  └───────────────── Center code
 └──────────────────── Class code
```

Example: `41126038` = Class 4, Center 1, Boy, year 26, serial 038.

---

## 2. Code tables

### Class code (1 digit)
| Class | Code |
|---|---|
| 04 | 4 |
| 05 | 5 |
| 06 | 6 |
| 07 | 7 |
| 08 | 8 |
| 09 | 9 |
| 10 | **1** (not 10, because the code is one digit) |

### Center code (1 digit)
| Center | Code |
|---|---|
| চট্টগ্রাম কলেজিয়েট স্কুল (Chattogram Collegiate School) | 1 |
| পতেঙ্গা উচ্চ বিদ্যালয় (Patenga High School) | 2 |
| সিলভার বেলস কিন্ডারগার্টেন এন্ড গার্লস হাই স্কুল (Silver Bells Kindergarten & Girls High School; also covers Halishahar & Agrabad) | 3 |

### Gender code (1 digit)
| Gender | Code |
|---|---|
| Boy | 1 |
| Girl | 2 |

---

## 3. Serial rule (most important)

- The serial counts up **separately for each (Class + Center + Gender) group**.
- There are 7 classes × 3 centers × 2 genders = **42 independent counters**.
- A new roll = the group's **latest serial + 1**, zero-padded to 3 digits.
- Never reuse a serial. Never skip a serial unless the user asks.
- If several students are added to the same group, give them consecutive serials.

---

## 4. Latest issued rolls (source of truth)

Last updated: 14/10/2026

| Class | Gender | Center 1 (Collegiate) | Center 2 (Patenga) | Center 3 (Silver Bells) |
|---|---|---|---|---|
| 4  | Boy  | 41126047 | 42126051 | 43126086 |
| 4  | Girl | 41226064 | 42226060 | 43226083 |
| 5  | Boy  | 51126059 | 52126054 | 53126069 |
| 5  | Girl | 51226070 | 52226072 | 53226122 |
| 6  | Boy  | 61126053 | 62126066 | 63126060 |
| 6  | Girl | 61226057 | 62226064 | 63226076 |
| 7  | Boy  | 71126053 | 72126066 | 73126049 |
| 7  | Girl | 71226067 | 72226058 | 73226063 |
| 8  | Boy  | 81126049 | 82126037 | 83126041 |
| 8  | Girl | 81226032 | 82226047 | 83226054 |
| 9  | Boy  | 91126013 | 92126033 | 93126019 |
| 9  | Girl | 91226018 | 92226055 | 93226043 |
| 10 | Boy  | 11126022 | 12126005 | 13126010 |
| 10 | Girl | 11226013 | 12226013 | 13226014 |

**After generating any new roll, update the matching cell above** with the new
latest roll and change the "Last updated" date (DD/MM/YYYY), so the next run
continues from the right number.

---

## 5. How to generate a roll (step by step)

1. Get the student's **class**, **center** and **gender**.
2. Convert each one with the code tables in section 2.
3. Find that group's latest roll in section 4 and take its last 3 digits.
4. Add 1 and pad the result to 3 digits.
5. Join: `class code + center code + gender code + "26" + serial`.
6. Check that the result is 8 digits and not already used.
7. Update section 4.

### Worked examples
| Student | Codes | Latest | New roll |
|---|---|---|---|
| Class 4, Center 1, Boy | 4,1,1 | 41126038 | **41126039** |
| Class 10, Center 2, Girl | 1,2,2 | 12226011 | **12226012** |
| Class 5, Center 3, Girl | 5,3,2 | 53226110 | **53226111** |
| 3 more Class 9, Center 1, Boy | 9,1,1 | 91126009 | **91126010, 91126011, 91126012** |

---

## 6. Rules and edge cases

- If the class, center or gender is missing or unclear, **ask** instead of guessing.
- Class 10 always uses class code **1**. A roll starting with 1 is never class 1.
- If a serial would go past 999, stop and tell the user.
- If the user gives a list of students, return a table:
  `Name | Class | Center | Gender | Roll`.
- To decode a roll, read it digit by digit with the tables above.
