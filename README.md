<h1>Purpose & Context:</h1> 
During my time at Legend Fleet Solutions, (Vehicle upfit manufacturing company) I created this <bold>VBA Macro</bold> to make entering records into the company's ISAH (ERP) database easier.

The company manufactures different <bold>parts</bold> for vehicles including flooring, insulated wall liners and ceilings, and other insulation. Each of these parts require a record in the database that links all the required files to produce the part. Files include production files that program manufacturing machines, product images, and drawings. Essentially anything required to re-produce the part from scratch. 

<h2>Part Numbers</h2>
Each variation of part is tracked with a specialized SKU called a Part Number that details its specific use ex, product type, vehicle type and year etc. Records in the ERP system are centered around part numbers.

<h2>Revision Numbers</h2>
Each of these part's is also designed by specialized engineers and at times, parts need to be revised and recreated due to issues like misfitting, wrong dimensions etc. These revisions are tracked by the revision number, and a part directory must have all the files with the same revision number, otherwise the part is created wrong and the customer purchasing the part will be incontent.

<h2>Manual Entry Process</h2>
Creating these records in the company's Enterprise Reporting and Planning system "ISAH", was a long and tedious process, with many redundant processes making manual processing time very long. Hence, I optimized and automated the process to reduce the time needed by nearly 95%, consolidating all the required files & file paths into a single excel sheet. The process could be fully automated but due to manual error in the creation of the original directories (differing file/folder names, incorrect titles etc), an element of manual processing was mandatory to ensure the entry process went smoothly.

<h2>Macro Functionality</h2>
-Traverses directory of record and extracts relevant data from specific folders.<br>
-Parses information into a format enterable in the ERP. <br>
-Quickly validates part folder structure and revision consistency. (In a manufacturing setting, ensuring that the correct revision is being used as well as a valid folder is paramount to ensuring seamless operations) <br>
-Flags any missing or empty folders crucial to the manufacturing process. <br>
-Trims the root of the file extention and replaces with one allowing file to be opened by other departments.



<h1>[INSTALLATION]</h1>

Follow the instructions below to ensure the ISAH Part Creation Assistant is trusted and able to run.
<ol type="1">
<li>Procedure: Ensure macros are enabled</li>
File > Options > Trust Center > Trust Center Settings > Macro Settings > Enable VBA Macros

![image](https://github.com/user-attachments/assets/0a1d1744-580c-4742-821a-c6f8783321b6)

<li>Save the ISAH Part Creation Assistant workbook to your downloads as "ISAH Part Creation Assistnat (version 1)", any other name will cause an error in the program. This was done to ensure that no irrelevant workbooks are damaged or meddled with.</li>
<li>Enable macros on the file
   -In your file explorer, right click on the File > Properties > Security Settings > Unblock </li>
   
![image](https://github.com/user-attachments/assets/a943ae4f-c016-4c15-b7dc-a371787540a2)


<li>Open the ISAH Part Creation Assistant Workbook</li>
</ol>

<h12>[USAGE]</h1>

<ol type="1">
<li>Click the run button at the top of the screen<br>
   
![image](https://github.com/user-attachments/assets/a69332e2-65f3-4e33-b1c6-9dccf7230470)
</li>
<li>File dialog will be opened, from there select the desired part folder. Note that each part folder follows the same template which this macro takes advantage of<br>
   -Example: 
   
   ![image](https://github.com/user-attachments/assets/d1196213-9c9e-42e2-b295-c0cc6c1ab0c0)

  Ensure that the Folder Name displayed is the folder you wish to evaluate / use. 

</li>
<li> Algorithm will traverse directory and pull information from the relevant folders:<br>
<br>
      
Valid Folder Structure looks like:
   
![image](https://github.com/user-attachments/assets/47631b87-98c8-4e28-b059-0d086ec41478)

If any of the folders do not exist, or if the folders exist but are empty, a flag will be raised on the sheet.<br>
Examples of possible outputs on the sheet:

![image](https://github.com/user-attachments/assets/8f11b883-152b-4629-ad8e-597dadb9329d)


   
</li>

<li>Revision numbers for the following folders are compared to ensure consistency: Shop Images, Product Images, and Drawing. <br>Examples of possible output include: <br>
<ul> 
<li>Inconsistent Rev Numbers (At least 1 not consistent or not found)

![image](https://github.com/user-attachments/assets/05e19610-d9e9-4113-a709-ce875dfc13f9)

</li>

<li>Consistent Rev Numbers (all consistent)

![image](https://github.com/user-attachments/assets/6c46f1d9-8a8e-4f19-aac1-ae83a8cd99e8)

</li>

<li> No Rev Numbers (Empty File)

![image](https://github.com/user-attachments/assets/64459489-bdb6-4690-96e5-3c88a7b1a379)

   
</li>
<li> Similar Rev Numbers (Off by a decimal point)

![image](https://github.com/user-attachments/assets/6fbfe4af-9365-483d-958c-2f48a7f97093)

</li>
</ul>
</li>
</ol>
