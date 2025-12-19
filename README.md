# Google Sheets Invoice Generator

A Google Apps Script that generates and displays an invoice summary from a "Report" sheet with a custom menu item. The script automatically determines the billing period from the "Payments" sheet and displays a formatted invoice in Persian/Farsi.

## Features

- Custom menu item "📜 Invoice" with "Show Invoice" submenu
- Automatically extracts pricing and team data from the Report sheet
- Reads billing period from Payments sheet to display subscription details
- Displays formatted invoice summary in a modal dialog with Vazirmatn font
- Supports Persian/Farsi text formatting with RTL direction
- Conditional display of team lines (only shows teams with at least one non-zero seat count)
- Currency formatting with dollar signs displayed correctly in RTL text

## Setup

1. Open your Google Sheet
2. Go to **Extensions** → **Apps Script**
3. Copy and paste the code from `invoice-maker.gs` into the script editor
4. Save the project (Ctrl+S or Cmd+S)
5. Refresh your Google Sheet
6. The "📜 Invoice" menu will appear in the menu bar

## Requirements

Your Google Sheet must have two sheets:

### 1. "Report" Sheet

#### Column Headers (Row 1)
- `Full` - Full seat price
- `Dev` - Developer seat price
- `Collab` - Collaboration seat price
- `Team` - Used for finding "Subtotal", "Prorated Costs", and "Total" rows
- `Name` - Team names (used for display)
- `Cost` - Cost values

#### Data Structure
- **Row 2**: Contains price values for Full, Dev, and Collab columns
- **Rows 3+**: Team data (name in "Name" column, seat counts in Full/Dev/Collab columns)
- **"Subtotal" row**: Contains totals in Team column (seat counts and cost subtotal)
- **"Prorated Costs" row**: Contains prorated cost value in Cost column
- **"Total" row**: Contains final total cost in Cost column

### 2. "Payments" Sheet (Optional)

#### Column Headers (Row 1)
- `Month` - Month name (e.g., "January", "February", etc.)

#### Data Structure
- Contains payment records with month values
- Script finds the last row with a Month value and determines the next billing period
- If this sheet doesn't exist or has no data, the subscription details line will be omitted from the invoice

## Usage

1. Click on **📜 Invoice** in the menu bar
2. Select **Show Invoice**
3. A modal dialog will display the invoice summary with:
   - Subscription details (month and period) - if Payments sheet has data
   - Price per seat for each type (with dollar signs)
   - Team breakdown with seat counts (only teams with at least one seat)
   - Total seat counts
   - Final cost calculation (with conditional prorated costs display)

## Output Format

The invoice displays in Persian/Farsi with RTL direction:

1. **Subscription Details** (if available):
   - "جزئیات اشتراک فیگما فاکتور ماه {month} (از {periodStart} تا {periodEnd}) به شرح زیر است:"

2. **Price Line**:
   - "قیمت هر سیت: فول ($X) — دولوپر ($Y) — کلب ($Z)"

3. **Team Lines** (only if team has at least one non-zero seat count):
   - "{TeamName} : X فول - Y دولوپر - Z کلب"
   - Only displays seat types that are greater than 0

4. **Total Seats**:
   - "کل سیت‌ها: X فول - Y دولوپر - Z کلب"

5. **Cost Summary**:
   - If no prorated costs: "مجموعا مبلغ نهایی $X"
   - If prorated costs exist: "مجموعا $X بعلاوه $Y هزینه سرشکن ماه قبل، مبلغ نهایی $Z"

## Technical Details

- **Font**: Vazirmatn (with sans-serif fallback)
- **Text Direction**: RTL (Right-to-Left) for Persian text
- **Currency Formatting**: Dollar signs are wrapped in LTR spans to ensure correct display in RTL text
- **Month Template**: Built-in template maps months to billing periods (30th of current month to 29th/30th of next month)

## Error Handling

The script includes error handling for:
- Missing "Report" sheet
- Missing required columns (Full, Dev, Collab, Team, Name, Cost)
- Missing "Subtotal" row
- Missing "Prorated Costs" or "Total" rows (defaults to 0)
- Missing "Payments" sheet (subscription details omitted)
- Missing "Month" column in Payments sheet (subscription details omitted)
- Empty Payments sheet (subscription details omitted)

