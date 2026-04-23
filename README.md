<p align="center">
  <img src="https://capsule-render.vercel.app/api?type=waving&color=0:0078D4,100:00C6FF&height=220&section=header&text=MRF%20Utility%20Excel%20Add-in&fontSize=32&fontColor=ffffff&animation=fadeIn&fontAlignY=35&desc=Stock%20Taking%20%7C%20Automation%20%7C%20SAP%20Integration&descAlignY=55&descSize=16" />
</p>

# 📗 MRF Utility Excel Add-In

[![Version](https://img.shields.io/badge/version-1.0.5-blue.svg "Download Excel Add-in v-1.0.5")](https://github.com/MRF-UTILITY/MRF/raw/main/StockTaking55-v-1.0.5.xlam)

Welcome to the **MRF Utility**—an advanced Excel add-in designed to optimize stock-taking and tallying processes. Whether you're managing stock data or automating workflows, this tool ensures precision and saves valuable time.

## 📑 Table of Contents
- [Key Features](#-key-features)
- [Prerequisites](#-prerequisites)
- [Installation Steps](#️-installation-steps)
- [How to Use](#how-to-use)
- [Updates & Version Control](#-updates--version-control)
- [Continuous Improvement](#-suggesting-enhancements-for-continuous-improvement)
- [Support](#-support)

---

## 🚀 Key Features

| Feature | Description | SAP Transaction |
|---|---|---|
| 📦&nbsp;**Stock&nbsp;Taking** | Easily manage stock-taking sheets by Product Groups and efficiently tally data for the INS Tube Flap | `ZNBPSTK`, `ZDL` |
| ✅&nbsp;**Set&nbsp;Checker** | Validate and cross-check set billings instantly with a single click | `ZIRN_VIEW_LINES` |
| 🔍&nbsp;**Find&nbsp;Error&nbsp;Invoice** | Retrieve invoice numbers and dates where proper set billing were missing | *(Run after completion of Set Checker)* |
| 📈&nbsp;**Pull&nbsp;Compliance** | Automate FDC-to-SOF workflows using predefined templates | *(Run on un-edited downloaded Pull file)* |

---

## 📋 Prerequisites

* **Microsoft Excel:** 2010 or later.

> [!IMPORTANT]
> **Trust Center Settings:** Before using this add-in, you must configure your Excel Trust Center settings to allow the add-in to work. *(One-Time Setup)*. [See the step-by-step setup guide](https://github.com/MRF-UTILITY/MRF/blob/main/Guide%20to%20Install%20Add-in%20file%20in%20Excel.pptx).

---

## 🛠️ Installation Steps

1. **Download:** Grab the latest add-in `.xlam` release from the [GitHub Releases](https://github.com/MRF-UTILITY/MRF/blob/main/StockTaking55-v-1.0.5.xlam) page.
2. **Unblock the File:** > **Important Windows Security Step:** > Locate the downloaded `.xlam` add-ins file, right-click it, and select **Properties**. At the bottom of the *General* tab under *Security*, check the **Unblock** box, and click **OK**.
3. **Open Excel:** Launch Microsoft Excel.
4. **Navigate to Add-ins:** Go to `File` > `Options` > `Add-ins`.
5. **Manage Add-ins:** At the bottom of the window, ensure *Manage* is set to **Excel Add-ins**, then click **Go...**.

6. **Browse and Install:** Click **Browse**, locate your unblocked `.xlam` add-in file, select it, and click **OK**.
7. **Access the Utility:** You will now see a dedicated **M R F** tab in your Excel Ribbon!

> [!NOTE]
> **Troubleshooting:** If you have followed all the above Installation steps properly and still the **M R F** ribbon tab does not appear, simply restart the Excel.

---
<a id="how-to-use"></a>
## ⚙️ How to Use

1. **Export Data from SAP:** Generate and download your spreadsheets from the relevant SAP transactions.
2. **Open the M R F Tab:** Navigate to the new custom tab in your Excel ribbon.

3. **Execute Workflows:**
   * **Stock Taking:** Run to process product-group-wise stock-taking sheets and to review the INS Tube Flap discrepancy in stock. Use SAP transactions (`ZNBPSTK`, `ZDL`). [See the Instruction Guide](https://github.com/MRF-UTILITY/MRF/blob/main/Stock%20Taking%20Tool%20Guide.docx).
     
   * **Set Checker:** Run the validation to instantly report on set billings. Use SAP transaction (`ZIRN_VIEW_LINES`). [See the Instruction Guide](https://github.com/MRF-UTILITY/MRF/blob/main/Set%20Checking%20Tool%20Guide.docx).
     
   * **Find Error Invoice:** Run to quickly review Invoice numbers and dates for Improper Set Billing. Best to run after the *Set Checker* tool process is completed.
     
   * **Pull Compliance:** Run to manage FDC to SOF workflows using T-Code `ME21N`, `MB1B`. Make sure the *`Source File.xlsx`* Excel workbook is available in the `FDC TO SOF` folder on your D Drive. *(Path: `D:\FDC TO SOF\SOURCE FILE.xlsx`)*.

> [!WARNING]
> ⚠️ ***Very Important Pull Compliance Step:*** Users must review and update the `Source File.xlsx` workbook located in `D:\FDC TO SOF\` themselves for any missing/Incorrect material codes where proper INS Tube/Flaps are not appearing. You must also add any new Tube-Type Tyre sizes to the list from time to time to ensure proper set billing from FDC to SOFs in IBTAs.  Keeping this file up to date is essential for accurate set billing.

---

## 🔄 Updates & Version Control

This add-in will automatically check against the GitHub repository to notify you if a newer version of the utility is available, ensuring you are always running the latest code and features.

---

## 💡 Suggesting Enhancements for Continuous Improvement

Excited to collaborate with contributors! To help enhance this tool & add more features, please follow these steps:

1. Fork this repository.
   ```bash
   git clone https://github.com/YOUR-USERNAME/MRF.git   
2. Create a new branch for your feature:
   ```bash
   git checkout -b feature/your-feature-name
3. Commit your changes:
   ```bash
   git commit -m "Add your feature"
4. Push the branch:
   ```bash
   git push origin feature/your-feature-name

---

## 💬 Support

Have questions, suggestions, or running into issues? Reach out via the [Issues](../../issues) section on GitHub. Happy to help!

