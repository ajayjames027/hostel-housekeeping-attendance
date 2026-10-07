# Hostel Housekeeping Attendance & Remuneration System

A complete automated attendance and payroll system tailored for hostel housekeeping staff.

---

## 📂 Included Files

1. **[`Hostel_Housekeeping_Attendance_and_Payroll.xlsx`](file:///c:/Users/ADMIN/Downloads/hostel_attendance/Hostel_Housekeeping_Attendance_and_Payroll.xlsx)**
   - **Sheet 1 (`Staff_Master`)**: Housekeeping staff directory, base monthly salary, daily wage rate, weekly off day, and payment/bank/UPI details.
   - **Sheet 2 (`Printable_Sign_Sheet`)**: Pre-formatted monthly physical register (Day 1 to 31) with Morning & Evening sign-in boxes, ready to print for physical signatures on notice board.
   - **Sheet 3 (`Attendance_Log`)**: Digital attendance ledger (`P` = Present, `HD` = Half Day, `WO` = Weekly Off, `PL` = Paid Leave, `A` = Absent). Automatically computes total payable days using dynamic Excel formulas.
   - **Sheet 4 (`Monthly_Remuneration`)**: Automated salary computation connected to attendance with columns for Basic Wage, Pro-rata Earned Pay, Overtime/Allowances, Deductions/Advance, Net Payable, and Disbursement Signature.

2. **[`Printable_Attendance_Register.html`](file:///c:/Users/ADMIN/Downloads/hostel_attendance/Printable_Attendance_Register.html)**
   - Ready-to-print **A4 Landscape** monthly sign-in register. Open in any web browser and press `Ctrl + P` to print.

3. **[`Printable_Remuneration_and_Payslips.html`](file:///c:/Users/ADMIN/Downloads/hostel_attendance/Printable_Remuneration_and_Payslips.html)**
   - Ready-to-print **A4 Portrait** monthly wage summary table and individual salary voucher receipts with staff acknowledgement signatures.

4. **[`generate_attendance_system.py`](file:///c:/Users/ADMIN/Downloads/hostel_attendance/generate_attendance_system.py)**
   - Python automation script to generate updated sheets for any future month or year and with updated staff names/salaries.

---

## 🧮 Remuneration Calculation Rules

$$\text{Daily Wage Rate} = \frac{\text{Monthly Base Wage}}{\text{Total Days in Month}}$$

$$\text{Payable Days} = \text{Present Days (P)} + 0.5 \times \text{Half Days (HD)} + \text{Weekly Offs (WO)} + \text{Paid Leaves (PL)}$$

$$\text{Earned Basic Wage} = \text{Daily Wage Rate} \times \text{Payable Days}$$

$$\text{Net Payable Wage} = \text{Earned Basic Wage} + \text{Overtime / Allowances} - \text{Advance / Deductions}$$

---

## 🚀 How to Customize or Generate for Future Months

To change the month, hostel name, or staff salaries, simply run:
```powershell
python generate_attendance_system.py
```
Or edit the values directly inside the [Excel workbook](file:///c:/Users/ADMIN/Downloads/hostel_attendance/Hostel_Housekeeping_Attendance_and_Payroll.xlsx).
