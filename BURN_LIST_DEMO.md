# Engineering Tools - Burn List Feature Demo

## 🔥 **NEW FEATURE: Create Burn List**

You can now see the **orange "Create Burn List"** button in the toolbar!

### **How It Works:**

1. **Engineering Workflow (First):**
   - Load Excel MRP data ✅ (You already tested this!)
   - Generate cutlist for jobs ✅ (You already tested this!)
   - Save job records (assigns XML files to jobs) ✅

2. **Manufacturing Workflow (New!):**
   - Click **"Create Burn List"** button (orange button in toolbar)
   - System finds jobs ready for burning (with XML files assigned)
   - User selects format:
     - **ERP format (.erp)** - for ERP systems
     - **WOL format (.wol)** - for Sigma cutting equipment
   - Files generated automatically in `C:\Scripts\EngineeringTools\xmlCutlist\`

### **What Files Look Like:**

**ERP Format (.erp):**
```xml
<ErpExchange>
    <Orders>
        <ErpOrder>
            <ImportType>NewOrder</ImportType>
            <OrderNumber>H1319-A</OrderNumber>
            <StartDate>2025-10-03</StartDate>
            <TargetDate>2025-10-10</TargetDate>
            <ProductionStrategy>MaterialAdministrationOrder</ProductionStrategy>
            <Automatic>True</Automatic>
            <PartNumber>12345-001</PartNumber>
            <Quantity>5</Quantity>
            <XmlFile>part_12345-001.xml</XmlFile>
        </ErpOrder>
    </Orders>
</ErpExchange>
```

**WOL Format (.wol):**
```
# WOL Burn List Generated: 10/3/2025 7:00:00 AM
H1319-A	12345-001	5	part_12345-001.xml	Sample Part Description
H1320-B	12346-002	3	part_12346-002.xml	Another Part Description
```

### **Key Features:**
- ✅ **Sequential Workflow** - Only jobs with engineering-assigned XML files can be burned
- ✅ **Dual Format Support** - ERP and WOL formats for different equipment
- ✅ **Automatic Backup** - Timestamped backup files created
- ✅ **Error Handling** - Validates engineering completion
- ✅ **User-Friendly** - Shows success message with file location option

### **Workflow Integration:**
```
Engineering → Manufacturing
    ↓             ↓
Cutlist    →  Burn List
(What to     (How to
 make)        make it)
```

The application is now running and you should see the orange **"Create Burn List"** button in the toolbar next to the green "Load from Excel" button!