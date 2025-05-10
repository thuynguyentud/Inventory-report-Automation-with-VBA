# Daily Inventory Report Automation with VBA

This is a VBA-powered Excel tool I developed for a project in 2022 in Vietnam to automate daily reporting tasks based on raw data exports from a Warehouse Management System (WMS) used in an FMCG (Fast-Moving Consumer Goods) warehouse.

## 📌 Background & Pain Points

In the original workflow, the WMS system allowed raw data exports, but these exports were not formatted or aggregated in a way that could be shared directly with customers or stakeholders.

Warehouse specialists had to:
- Download multiple sheets from different WMS reports daily
- Manually filter and calculate data from different sheets based on transaction type and logic
- Format the numbers and structure the data to match a required report layout
- Export the report as an Excel file
- Manually write and send an email with the report attached

Despite being repetitive, the complexity of data conditions meant the task required **concentration and took ~30 minutes each day**.

---

## ✅ Improvement with VBA:

Using VBA embedded in an Excel template, the tool automates:
- Data aggregation across multiple sheets
- Conditional filtering and column-specific calculations
- Standardized number formatting
- Generating a final report sheet
- Exporting the report as an Excel file
- Preparing an Outlook email with the report attached

The entire process is reduced from **~30 minutes to just 5 minutes**, achieving a **~83% productivity gain**.

## 🔐 Data Privacy

All data used is **synthetically generated** to reflect the typical WMS reports in FMCG warehouses. No confidential business data is included.

## 📁 File Structure

- `Inventory daily report.xlsm` – Main Excel file with raw data, report template, and embedded VBA macros
- **Sheet1 & Sheet2**: Simulated WMS raw data views, with Sheet 1 containing raw data of the inventory quantity and statuses, and Sheet 2 containing a transaction-level view showing inventory activities.
- **Sheet4**: Formatted report layout (final output)
  
## 🔧 Main VBA Tasks

- Lookup and sum data by `SKU ID` under specific transaction types (e.g., RECEIVING, SHIPPING)
- Cross-sheet calculations
- Formatting output columns (thousands separator, no decimals)
- Exporting the report as an `.xlsx` file
- Preparing and displaying an Outlook email draft with a recipient list and report attached

## 📈 Outcome

- Reduced mental load and human error
- Improved consistency and formatting
- Reusable template for similar reporting cases

## 📬 How to Use
1. Open the file `Inventory daily report.xlsm`, view sheet 4 `Report_template`
2. Click Alt + F8 and Run `RunAll` from the Macro tab
   
  ![](https://github.com/user-attachments/assets/c84dfb8f-f788-49d4-bf64-d92e255e2e12)

Excel returns a Notification "All report columns have been updated and formatted", with all blank columns in the report filled with aggregated data.

   ![](https://github.com/user-attachments/assets/bbceaeba-6f8d-42ad-9e45-1169fed06bb9)

3. Repeat step 3 and Run `SendReportEmail` from the Macro tab. Excel returns a Notification "Report saved and email prepared".
   
   ![](https://github.com/user-attachments/assets/38c1a426-1583-4270-99de-b3bc6442f2fa)

At this step, an Excel report file is automatically downloaded and the Email view is opened with the attached report file, recipients list, and email content.

![](https://github.com/user-attachments/assets/485f3148-4d2c-440c-8ccc-d0d2a16cb7f4)

To view the VBA code, click Alt + F11 to view. Here is the picture of the VBA code I created.

![image](https://github.com/user-attachments/assets/56d3f460-ecd4-470c-a5ae-7aa6e0455d38)

---

## 📄 License

This project is licensed under the **MIT License**.

You are free to use, copy, modify, merge, publish, and distribute this code, provided that proper attribution is given.

> © 2022 Thu Thuy Nguyen.  
> If you use or adapt this project, please cite the original repository and author.

---

