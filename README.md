
### **Overview**

The automation process streamlines the export, transformation, and delivery of costing data from the PLM system to customers. Each customer requires the data in specific formats, necessitating custom processing and formatting. This solution enhances efficiency, minimizes errors, and ensures timely delivery of customer-specific data.

### **Process Workflow**
![image](https://github.com/user-attachments/assets/30687b84-e457-4385-8aac-4ce69117c45e)

#### **1. Data Export from PLM**  
- Costing data is exported from the PLM web system as an Excel file.  
- The exported file contains raw data required for processing.

#### **2. File Storage**  
- The exported Excel file is uploaded to a designated folder in **SharePoint** for centralized access and management.  

#### **3. Data Transformation Using Power Query**  
- **Power Query** is utilized for extracting and transforming the raw data.
- Transformed data is loaded back into an Excel workbook for further processing.

#### **4. Customer-Specific Formatting**  
- The processed data is integrated into **customer-specific templates**:  
  - Relevant data fields are mapped to the appropriate cells.  
  - Predefined calculations are applied using formulas within the Excel template.  

#### **5. Advanced Formatting**  
- Certain customer formats require advanced scripting for formatting:  
  - scripting is employed for tasks such as:  
    - Conditional formatting.  
    - Custom layouts and designs.  
    - Adjusting font styles, colors, and other visual elements.  
![image](https://github.com/user-attachments/assets/cce4979e-f50b-453f-939f-19c7465e711a)

#### **6. Finalization and Delivery**  
- The finalized customer-specific files are reviewed for accuracy.  
- Files are delivered to customers via email or uploaded to designated platforms.

---

### **Solution Features**

- **Automated Workflow**: Reduces manual effort through seamless integration of Power Query and SharePoint.  
- **Customizable Templates**: Ensures each customer receives data in their required format.  
- **Scalability**: The solution can accommodate additional customers or format changes with minimal effort.  
- **Error Reduction**: Automated processes minimize manual errors.  
- **Time Efficiency**: Rapid processing and delivery improve customer satisfaction.  

---

### **Technologies Used**

- **PLM Web System**: Data source.  
- **SharePoint**: Centralized file storage and access.  
- **Power Query**: Data extraction and transformation tool.  
- **Excel**: Data loading, calculation, and formatting.  
- **Scripting**: VBA and/or Power Automate for advanced formatting.

---

