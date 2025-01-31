# 💡 ISAH Part Creation Assistant  

## ℹ️ Project Overview and Important Context  
During my time at Legend Fleet Solutions, (a vehicle upfit manufacturing company) I created this **VBA Macro** to streamline the process of entering records into the company's ISAH ERP database.

## 🚚 About Legend Fleet Solutions  
Legend Fleet Solutions manufactures various **vehicle parts**, including:

- **Flooring, insulated wall liners, ceilings, and other insulation components.**
- **Each part requires a corresponding record in the ERP system, linking all necessary files for manufacturing.**
- **These files include:**
  - Production files for manufacturing machines.
  - Product images and technical drawings.
  - Any documentation required to reproduce the part from scratch.

## 🔢 Part Numbers  
Each part variation is tracked using a **specialized SKU** called a **Part Number**, which contains important details such as:

- Product Type.
- Vehicle Type.
- Model Year.
- Other relevant specifications.

All records in the **ERP System** are centered around **Part Numbers**.

## 🔄 Revision Numbers  
Parts are designed by **specialized engineers**, and at times, revisions are necessary due to:

- Misfitting.
- Incorrect Dimensions.
- Other manufacturing issues.

These revisions are tracked using a **Revision Number**.  
For a directory to be accurate, it must contain files **with the same revision numbers**. Any inconsistencies could cause incorrect production, resulting in **customer dissatisfaction and operational inefficiencies**.

---

## 🏗️ Manual Entry & Search Process (Before Automated Creation)  
Prior to creating this operation, creating records in the ISAH ERP system was a **long and tedious** process.

- **Redundant steps** and **manual file searching** made this process inefficient.
- Manual errors in **file naming, folder structures, and missing files** caused frequent delays.
- The entire process was **very time-consuming**.

---

## ⚡ Impact of My Macro  
To optimize and automate this process, I developed a **VBA Macro** that:

- **Reduced processing time by 95% by automatically pulling data from part directories.**
- **Consolidated all required files and file paths into a single Excel sheet.**
- **Reduced inconsistencies by 95% by cross-referencing revision numbers of all files within part directory.**
- **Saved hours of time processing records manually.**

While full automation into the ERP was possible, an element of **manual review was necessary** due to inconsistencies in original **file and folder names**.

---

## 🛠 Installation  
Follow the instructions below to ensure the ISAH Part Creation Assistant is trusted and able to run.

1. **Enable Macros in Excel:**  
    - Navigate to **File > Options > Trust Center > Trust Center Settings > Macro Settings**  
    - Enable VBA Macros.
     ![image](https://github.com/user-attachments/assets/0a1d1744-580c-4742-821a-c6f8783321b6)

2. **Save the ISAH Part Creation Assistant workbook to your downloads as "ISAH Part Creation Assistant (version 1)",**
    - Any other name will cause an error in the program. 

3. **Unblock the file:**
    - Right-click on the file  
    - Go to **Properties > Security Settings**  
    - Click **Unblock**  
![image](https://github.com/user-attachments/assets/a943ae4f-c016-4c15-b7dc-a371787540a2)
---

## 🚀 Usage
1. Open the ISAH Part Creation Assistant Workbook

2. **Click the run button at the top of the screen:**
   ![Run Button](https://github.com/user-attachments/assets/a69332e2-65f3-4e33-b1c6-9dccf7230470)
   

3. **Select the desired part folder.**  
- Example:  
![Folder Selection](https://github.com/user-attachments/assets/d1196213-9c9e-42e2-b295-c0cc6c1ab0c0)

4. **Algorithm will traverse the directory and pull relevant data.**  
- Valid folder structure example:  
![Valid Structure](https://github.com/user-attachments/assets/47631b87-98c8-4e28-b059-0d086ec41478)

5. **If any folders are missing or empty, a flag will be raised.**
-Example output:
![image](https://github.com/user-attachments/assets/8f11b883-152b-4629-ad8e-597dadb9329d)


## 🔍 Revision Number Validation

- If revision numbers acrross folders **are inconsistent** (at least 1 not consistent or not found), an alert is shown.

![image](https://github.com/user-attachments/assets/05e19610-d9e9-4113-a709-ce875dfc13f9)

- If **all revision numbers are consistent**, a confirmation is displayed.

![image](https://github.com/user-attachments/assets/6c46f1d9-8a8e-4f19-aac1-ae83a8cd99e8)


- If revision numbers are **slightly off**, a warning is displayed.
![image](https://github.com/user-attachments/assets/6fbfe4af-9365-483d-958c-2f48a7f97093)


