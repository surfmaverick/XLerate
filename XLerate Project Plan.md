# XLerate Module Update Checklist

## 📋 Current Modules Analysis

### Core Object Modules
- [ ] **ThisWorkbook.cls** - ⚠️ CRITICAL UPDATE REQUIRED
  - Current: Basic shortcuts (`Ctrl+Shift`)
  - Update: Add Macabacus-style shortcuts (`Ctrl+Alt+Shift`)
  - Status: Complete replacement needed

### Form Modules (.frm files)
- [ ] **frmNumberSettings.frm** - ✅ Minor updates
- [ ] **frmCellSettings.frm** - ✅ Minor updates  
- [ ] **frmDateSettings.frm** - ✅ Minor updates
- [ ] **frmTextStyle.frm** - ✅ Minor updates
- [ ] **frmAutoColor.frm** - ✅ Minor updates
- [ ] **frmErrorSettings.frm** - ✅ Minor updates
- [ ] **frmSettingsManager.frm** - 🔧 Add new currency settings panel
- [ ] **frmPrecedents.frm** - ✅ No changes needed
- [ ] **frmDependents.frm** - ✅ No changes needed

### Standard Modules (.bas files)
- [ ] **ModNumberFormat.bas** - ✅ Minor updates
- [ ] **ModCellFormat.bas** - ✅ Minor updates
- [ ] **ModDateFormat.bas** - ✅ Minor updates
- [ ] **ModTextStyle.bas** - ✅ Minor updates
- [ ] **ModGlobalSettings.bas** - ✅ Minor updates
- [ ] **ModSettings.bas** - ✅ No changes needed
- [ ] **ModCAGR.bas** - ✅ No changes needed
- [ ] **ModFormatReset.bas** - ✅ No changes needed
- [ ] **ModErrorWrap.bas** - ✅ No changes needed
- [ ] **ModSmartFillRight.bas** - ✅ No changes needed
- [ ] **ModSwitchSign.bas** - ✅ No changes needed
- [ ] **FormulaConsistency.bas** - ✅ No changes needed
- [ ] **AutoColorModule.bas** - ✅ No changes needed
- [ ] **TraceUtils.bas** - 🔧 Enhance with new features
- [ ] **RibbonCallbacks.bas** - 🔧 Add new callback functions

### New Modules to Add
- [ ] **ModFastFillDown.bas** - 🆕 NEW MODULE
- [ ] **ModCurrencyCycling.bas** - 🆕 NEW MODULE
- [ ] **ModAdvancedAuditing.bas** - 🆕 NEW MODULE (optional)
- [ ] **ModAdvancedPaste.bas** - 🆕 NEW MODULE (optional)

### Class Modules (.cls files)
- [ ] **clsFormatType.cls** - ✅ No changes needed
- [ ] **clsCellFormatType.cls** - ✅ No changes needed
- [ ] **clsTextStyleType.cls** - ✅ No changes needed
- [ ] **clsUISettings.cls** - ✅ No changes needed
- [ ] **clsListBoxHandler.cls** - ✅ No changes needed
- [ ] **clsDynamicButtonHandler.cls** - ✅ No changes needed

### Ribbon/UI Files
- [ ] **customUI14.xml** - 🔧 Add new ribbon buttons

---

## 🚀 Implementation Phases

### Phase 1: Core Shortcuts (CRITICAL - Do First)
1. [ ] Update **ThisWorkbook.cls** with new shortcut system
2. [ ] Add **ModFastFillDown.bas** 
3. [ ] Add **ModCurrencyCycling.bas**
4. [ ] Test basic functionality

### Phase 2: Enhanced Features  
1. [ ] Update **RibbonCallbacks.bas** with new functions
2. [ ] Update **customUI14.xml** with new ribbon buttons
3. [ ] Enhance **TraceUtils.bas** with additional features

### Phase 3: Settings Integration
1. [ ] Update **frmSettingsManager.frm** to include currency settings
2. [ ] Test all settings panels work correctly

---

## ⚠️ Priority Levels

### 🔴 CRITICAL (Must Do)
- **ThisWorkbook.cls** - Breaks functionality if not updated
- **ModFastFillDown.bas** - Core new feature
- **ModCurrencyCycling.bas** - Core new feature

### 🟡 IMPORTANT (Should Do)  
- **RibbonCallbacks.bas** - New functionality won't work from ribbon
- **customUI14.xml** - Users won't see new features in ribbon

### 🟢 OPTIONAL (Nice to Have)
- **TraceUtils.bas** enhancements
- **frmSettingsManager.frm** currency panel
- Additional new modules

---

## 📁 File Structure Reference

```
src/
├── objects/
│   └── ThisWorkbook.cls ⚠️ CRITICAL UPDATE
├── forms/
│   ├── frmNumberSettings.frm
│   ├── frmCellSettings.frm  
│   ├── frmDateSettings.frm
│   ├── frmTextStyle.frm
│   ├── frmAutoColor.frm
│   ├── frmErrorSettings.frm
│   ├── frmSettingsManager.frm 🔧 UPDATE
│   ├── frmPrecedents.frm
│   └── frmDependents.frm
├── modules/
│   ├── ModNumberFormat.bas
│   ├── ModCellFormat.bas
│   ├── ModDateFormat.bas
│   ├── ModTextStyle.bas
│   ├── ModGlobalSettings.bas
│   ├── ModSettings.bas
│   ├── ModCAGR.bas
│   ├── ModFormatReset.bas
│   ├── ModErrorWrap.bas
│   ├── ModSmartFillRight.bas
│   ├── ModSwitchSign.bas
│   ├── FormulaConsistency.bas
│   ├── AutoColorModule.bas
│   ├── TraceUtils.bas 🔧 ENHANCE
│   ├── RibbonCallbacks.bas 🔧 UPDATE
│   ├── ModFastFillDown.bas 🆕 NEW
│   └── ModCurrencyCycling.bas 🆕 NEW
├── class modules/
│   ├── clsFormatType.cls
│   ├── clsCellFormatType.cls
│   ├── clsTextStyleType.cls
│   ├── clsUISettings.cls
│   ├── clsListBoxHandler.cls
│   └── clsDynamicButtonHandler.cls
└── ribbon/
    └── customUI14.xml 🔧 UPDATE
```

---

## 🔍 Quick Status Legend
- ✅ No changes needed
- 🔧 Minor updates required  
- ⚠️ Critical update required
- 🆕 New module to add
- 🔴 High priority
- 🟡 Medium priority  
- 🟢 Low priority