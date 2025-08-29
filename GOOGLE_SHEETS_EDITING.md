# Google Sheets Orders Editing Feature

This feature allows you to edit **ANY** fields in the Google Sheets **Orders** worksheet and **delete entire rows**, with all changes being saved back to your main BrickLink order files (XML/CSV).

## How It Works

When the tool runs, it:
1. Reads any existing user edits and deletions from the Orders worksheet
2. Saves edited data back to your source XML/CSV files
3. Removes deleted orders/items from your source XML/CSV files 
4. Processes the updated order data from your XML/CSV files
5. Updates the worksheet with the current data (preserving any remaining edits)

## Editable Fields

You can now edit **ALL** fields in the Orders worksheet, and your changes will be preserved:

- **Order Level Fields:**
  - **Order ID** - BrickLink order identifier
  - **Seller** - Seller name/username
  - **Order Date** - Date the order was placed
  - **Shipping** - Shipping costs for orders
  - **Add Chrg 1** - Additional charges (fees, taxes, etc.)
  - **Order Total** - Total amount paid for items
  - **Base Grand Total** - Total including fees
  - **Total Lots** - Number of different item types in the order
  - **Total Items** - Total quantity of all items in the order
  - **Tracking No** - Shipping tracking numbers

- **Item Level Fields:**
  - **Condition** - Item condition (N=New, U=Used, etc.)
  - **Item Number** - BrickLink item/part number
  - **Item Description** - Description of the item
  - **Color** - Color name
  - **Qty** - Quantity ordered
  - **Each** - Price per item
  - **Total** - Total price for this item

## New Features

### Row Deletion
- **Delete entire orders**: Delete the order header row to remove the entire order and all its items
- **Delete individual items**: Delete item rows to remove specific items from an order
- Deleted rows are permanently removed from your source XML/CSV files

### Full Field Editing
- Edit any field in the sheet - all changes are saved back to source files
- No longer limited to just shipping, tracking, and a few other fields
- Changes are applied to both XML and CSV files automatically

## Usage Instructions

1. **Run the tool normally** to populate the Orders worksheet with your BrickLink order data
2. **Edit any fields or delete rows** directly in Google Sheets:
   - Open your "Minifig Profit Tool" spreadsheet
   - Go to the "Orders" tab
   - Edit any field in any column - all changes will be saved back to your source files
   - Delete entire rows to permanently remove orders or items from your source files
   - Save your changes in Google Sheets
3. **Run the tool again** - your edits and deletions will be applied to your source files and preserved in the sheet

## Example Workflow

```
Initial state (from BrickLink data):
Order ID: 123456, Seller: TestSeller, Condition: N, Qty: 10, Each: $2.50

User edits in Google Sheets:
Order ID: 123456, Seller: EditedSeller, Condition: U, Qty: 12, Each: $3.00

User deletes an item row for part 3002 from the same order

After next tool run:
- Source XML/CSV files updated with: Seller: EditedSeller, Condition: U, Qty: 12, Each: $3.00
- Part 3002 permanently removed from source files
- Google Sheet reflects the updated data with edits preserved
```

## Important Notes

- **All fields editable**: You can now edit ANY field in the sheet, not just the previously limited set
- **Permanent changes**: Edits and deletions are saved back to your XML/CSV source files
- **Row deletion**: Deleting rows permanently removes that data from your source files
- **No validation**: The tool doesn't validate your edits - ensure data accuracy (correct formats, valid numbers, etc.)
- **Order identification**: Edits are matched to specific orders and items by Order ID and Item Number
- **Backup recommended**: Always keep backups of important data, especially since changes are now permanent

## Troubleshooting

- **Edits not preserved**: Make sure your Google Sheets file is saved before running the tool again
- **Formatting issues**: Numeric fields should contain valid numbers (e.g., "5.99" not "$5.99" - the tool will format currency display)
- **Deleted data reappears**: Ensure you're not re-importing the same data from duplicate source files
- **Permission errors**: Ensure the service account has proper access to your Google Sheets
- **Data corruption**: Always maintain backups since changes are now permanent in source files

## Technical Details

The enhanced feature works by:
1. Calling `read_orders_sheet_edits()` to capture ALL user data from the worksheet
2. Using `detect_deleted_orders()` to identify rows deleted by the user
3. Calling `remove_deleted_orders_from_files()` to permanently delete removed entries from source files
4. Using `save_edits_to_files()` to write edited data back to XML/CSV source files
5. Reloading fresh data from the updated source files
6. Using `update_orders_sheet()` to update the worksheet with current data

This provides full bi-directional synchronization between Google Sheets and your BrickLink order files.