<h1>💡 ISAH Part Creation Assistant </h1> 

<h2>ℹ️ Project Overview and Important Context</h2>
During my time at Legend Fleet Solutions (a vehicle upfit manufacturing company), I created this <bold>VBA Macro</bold> to streamline the process of entering records into the company's ISAH ERP database. 


<br><br>

<h3>🚚 About Legend Fleet Solutions:</h3>
Legend Fleet Solutions manufacturers various <bold>vehicle parts</bold>, including:<br>
<ul>
   <li>
      Flooring, insulated wall liners, ceilings, and other insulation components.
   </li>
   <li>
      Each part requires a corresponding <bold>record in the ERP system</bold>, linking all necessary files for manufacturing.
   </li>
   <li>
      These files include:<br>
      <ul>
         <li>Production files for manufacturing machines.</li>
         <li>Product images and technical drawings.</li>
         <li>Any documentation required to reproduce the part from scratch.</li>
      </ul>
   </li>
</ul>

<h3>🔢 Part Numbers</h3>
Each part variation is tracked using a <bold>specialized SKU</bold> called a <bold>Part Number,</bold>bold> which contains important details such as:<br><br>
<ul>
   <li>Product Type.</li>
   <li>Vehicle Type.</li>
   <li>Model Year.</li>
   <li>Other relevant specifications.</li>
   All records in the <bold>ERP System</bold> are centered around <bold>Part Numbers</bold>
</ul>


<h3>🔄 Revision Numbers</h3>
Parts are designed by <bold>specialized engineers,</bold> and at times, revisions are necesssary due to: <br><br>
<ul>
   <li>Misfitting.</li>
   <li>Incorrect Dimensions.</li>
   <li>Other manufacturing issues.</li>
</ul>
These revisions are tracked using a <bold>Revision Number</bold>
For a directory to be accurate, it must contain files <bold>with the same revision numbers</bold>. Any inconsistencies could cause incorrect production, resulting in <bold>customer dissatisfaction and operational inefficiencies</bold>.


<h2>🏗️  Manual Entry & Search Process (Before Automated Creation)</h2>
Prior to creating this operation, creating records in the ISAH ERP system was a <bold>long and tedious</bold> process. <br><br>
<ul>
   <li><bold>Redundant steps</bold> and <bold>manual file searching</bold> made this process inefficient.</li>
   <li>Manual errors in <bold>file naming, folder structures, and missing files</bold> caused frequent delays.</li>
   <li>The entire process was <bold>very tiem consuming</bold></li>
</ul>
<hr>
<h2>⚡ Impact of my Macro</h2>
To optimize and automate this process, I developed a <bold>VBA Macro</bold> that:<br><br>
<ul>
   <li>
      <bold>Reduced processing time by 95% by automatically pulling data from Part directories.</bold>
   </li>
   <li>
      <bold>Consolidated all required files and file paths into a single excel sheet</bold>
   </li>
   <li><bold>Reduced inconsistencies by 95% by cross referencing revision numbers of all files within part directory</bold></li>
   <li>S<bold>aved hours of time processing records manually</bold></li>
</ul>
While full automation into the ERP was possible, an element of <bold>manual review was necessary</bold> due to inconsistencies in original <bold>file and folder names</bold>

<h3></h3>
Creating these records in the company's Enterprise Reporting and Planning system "ISAH", was a long and tedious process, with many redundant processes making manual processing time very long. Hence, I optimized and automated the process to reduce the time needed by nearly 95%, consolidating all the required files & file paths into a single excel sheet. The process could be fully automated but due to manual error in the creation of the original directories (differing file/folder names, incorrect titles etc), an element of manual processing was mandatory to ensure the entry process went smoothly.

<h2>Macro Functionality</h2>
-Traverses directory of record and extracts relevant data from specific folders.<br>
-Parses information into a format enterable in the ERP. <br>
-Quickly validates part folder structure and revision consistency. (In a manufacturing setting, ensuring that the correct revision is being used as well as a valid folder is paramount to ensuring seamless operations) <br>
-Flags any missing or empty folders crucial to the manufacturing process. <br>
-Trims the root of the file extention and replaces with one allowing file to be opened by other departments.



<h1>Installation</h1>

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

<h1>Usage</h1>

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
