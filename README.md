# 📗 MRF Utility Excel Add-In

![Version](https://img.shields.io/badge/version-1.0.5-blue.svg)
![Platform](https://img.shields.io/badge/platform-Windows%20Excel-lightgray.svg)

Welcome to the **MRF Utility**—an advanced Excel add-in designed to optimize stock-taking and tallying processes. Whether you're managing stock data or automating workflows, this tool ensures precision and saves valuable time.

## 📑 Table of Contents
- [Key Features](#-key-features)
- [Prerequisites](#-prerequisites)
- [Installation Steps](#️-installation-steps)
- [How to Use](#-how-to-use)
- [Updates & Version Control](#-updates--version-control)
- [Contributing](#-contributing)
- [Support](#-support)

---

## 🚀 Key Features

* **Stock Taking Program:** Easily manage stock-taking sheets by Product Groups and efficiently tally data for the INS Tube Flap.
* **Customizable Reports:** Export data directly from SAP transactions (`ZNBPSTK`, `ZDL`) and process it seamlessly in Excel.
* **Set Checker Tool:** Validate and cross-check set billings with a single click for accurate reporting. (`ZIRN_VIEW_LINES`)
* **Find Error Invoice:** Quickly retrieve the Invoice numbers and dates where proper set billing were missing. *(Best to run immediately after the **Set Checker tool**).*
* **Pull Compliance Tool:** Automate FDC to SOF workflows using predefined templates to streamline data handling.

---

## 📋 Prerequisites

Before installing, ensure your system meets the following requirements:
* **Microsoft Excel:** 2010 or later.
* **<span style="color:red;">Trust Center Settings:</span>** Must be enabled in your Excel Trust Center (one-time setup). ([See the step-by-step setup guide](https://github.com/MRF-UTILITY/MRF/blob/main/Guide%20to%20Install%20Add-in%20file%20in%20Excel.pptx)).

---

## 🛠️ Installation Steps

1.  **Download:** Grab the latest `.xlam` release from the [GitHub Releases](https://github.com/MRF-UTILITY/MRF/blob/main/StockTaking55-v-1.0.5.xlam) page.
2.  **Unblock the File:** > ⚠️ **Important Windows Security Step:** > Locate the downloaded `.xlam` file, right-click it, and select **Properties**. At the bottom of the *General* tab under *Security*, check the **Unblock** box, and click **OK**.
3.  **Open Excel:** Launch Microsoft Excel.
4.  **Navigate to Add-ins:** Go to `File` > `Options` > `Add-ins`.
5.  **Manage Add-ins:** At the bottom of the window, ensure *Manage* is set to **Excel Add-ins**, then click **Go...**.
6.  **Browse and Install:** Click **Browse**, locate your unblocked `.xlam` file, select it, and click **OK**.
7.  **Access the Utility:** You will now see a dedicated **M R F** tab in your Excel Ribbon!

---

## ⚙️ How to Use

1.  **Export Data from SAP:** Generate and download your spreadsheets from the relevant SAP transactions.
2.  **Open the M R F Tab:** Navigate to the new custom tab in your Excel ribbon.
3.  **Execute Workflows:**
    * *Stock Taking:* Run to process product group wise stock taking sheets and to review the INS Tube Flap descripency in stock. Use SAP transaction (`ZNBPSTK`, `ZDL`). ([See the step-by-step Instruction File](https://github.com/MRF-UTILITY/MRF/blob/main/INS%20Tube%20Flap%20Stock%20aking%20Tool%20Guide.docx)).
    * *Set Checker:* Run the validation to instantly report on set billings. Use SAP transaction (`ZIRN_VIEW_LINES`). ([See the step-by-step Instruction File](https://github.com/MRF-UTILITY/MRF/blob/main/Set%20Checking%20Tool%20Guide.docx)).
    * *Pull Compliance:* Run to manage FDC to SOF workflows. (Make sure *Source File* excel workbook available in FDC TO SOF folder in D Drive. Path should be D:\FDC TO SOF\SOURCE FILE.xlsx). **VERY IMPORTANT**--User must review & update this Source file excel workbook themselves for any missing material codes where proper INS Tube/Flaps are not appearing & also for adding any new Tube-Type tyre in the list from time to time to ensuring proper set billing in IBTA's to SOF's from their FDC.     
    * *Find Error Invoice:* Run to quickly review Invoice numbers & dates for Improper Set Billing. Need to run only after Set Checker tool process completed.

---

## 🔄 Updates & Version Control

This add-in features built-in version checking. It will automatically check against the GitHub repository to notify you if a newer version of the utility is available, ensuring you are always running the latest code and features.

---

## 💡 Suggesting Enhancements for Continious Improvement

Excited to collaborate with contributors! To help enhance this tool & add more features, please follow these steps:

1. Fork this repository.
   
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


   
