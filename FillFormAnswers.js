async function fillFormAnswers(rowNumber) {
  // rowNumber = 2;
  const timezone = "GMT";
  const format = "dd-MM-yyyy";
  let formattedDate;

  const sheet = SpreadsheetApp.openById(
    RESPONSES_SPREADSHEET_ID
  ).getSheetByName(RESPONSES_SHEET_NAME);
  const headerRowValues = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()[0];
  const sheetValues = sheet.getDataRange().getValues();
  // const body = doc.getBody();
  const timestampColumnNo = headerRowValues.indexOf(TIMESTAMP);
  const auditDateColumnNo = headerRowValues.indexOf(AUDIT_DATE);
  const auditorNameColumnNo = headerRowValues.indexOf(AUDITOR_NAME);
  const warehouseCodeColumnNo = headerRowValues.indexOf(WAREHOUSE_CODE);
  const commodityHealthColumnNo = headerRowValues.indexOf(COMMODITY_HEALTH);
  const commoditiesAvailableColumnNo = headerRowValues.indexOf(
    COMMODITIES_AVAILABLE
  );
  const uploadPhotoCommodityColumnNo = headerRowValues.indexOf(
    UPLOAD_PHOTO_COMMODITY
  );
  const infestationDeteriorationNoticedColumnNo = headerRowValues.indexOf(
    INFESTATION_DETERIORATION_NOTICED
  );
  const infestationDeteriorationRemarksColumnNo = headerRowValues.indexOf(
    INFESTATION_DETERIORATION_REMARKS
  );
  const uploadPhotoInfestationDeteriorationColumnNo = headerRowValues.indexOf(
    UPLOAD_PHOTO_INFESTATION_DETERIORATION
  );
  const fumigationRequiredColumnNo =
    headerRowValues.indexOf(FUMIGATION_REQUIRED);
  const fumigationRequiredRemarksColumnNo = headerRowValues.indexOf(
    FUMIGATION_REQUIRED_REMARKS
  );
  const uploadPhotoFumigationColumnNo = headerRowValues.indexOf(
    UPLOAD_PHOTO_FUMIGATION
  );
  const dunnageAvailableColumnNo = headerRowValues.indexOf(DUNNAGE_AVAILABLE);
  const dunnageAvailableRemarksColumnNo = headerRowValues.indexOf(
    DUNNAGE_AVAILABLE_REMARKS
  );
  const uploadPhotoDunnageColumnNo =
    headerRowValues.indexOf(UPLOAD_PHOTO_DUNNAGE);
  const stockKeptCountableColumnNo =
    headerRowValues.indexOf(STOCK_KEPT_COUNTABLE);
  const stockKeptCountableRemarksColumnNo = headerRowValues.indexOf(
    STOCK_KEPT_COUNTABLE_REMARKS
  );
  const uploadPhotoStockKeptCountableColumnNo = headerRowValues.indexOf(
    UPLOAD_PHOTO_STOCK_KEPT_COUNTABLE
  );
  const weighingScaleAvailableColumnNo = headerRowValues.indexOf(
    WEIGHING_SCALE_AVAILABLE
  );
  const weighingScaleAvailableRemarksColumnNo = headerRowValues.indexOf(
    WEIGHING_SCALE_AVAILABLE_REMARKS
  );
  const uploadPhotoWeighingScaleColumnNo = headerRowValues.indexOf(
    UPLOAD_PHOTO_WEIGHING_SCALE
  );
  const hygieneCleanlinessMaintainedColumnNo = headerRowValues.indexOf(
    HYGIENE_CLEANLINESS_MAINTAINED
  );
  const hygieneCleanlinessMaintainedRemarksColumnNo = headerRowValues.indexOf(
    HYGIENE_CLEANLINESS_MAINTAINED_REMARKS
  );
  const uploadPhotoHygieneCleanlinessColumnNo = headerRowValues.indexOf(
    UPLOAD_PHOTO_HYGIENE_CLEANLINESS
  );
  const bankFundedCommoditiesStoredColumnNo = headerRowValues.indexOf(
    BANK_FUNDED_COMMODITIES_STORED
  );
  const bankFundedCommoditiesDetailsColumnNo = headerRowValues.indexOf(
    BANK_FUNDED_COMMODITIES_DETAILS
  );
  const uploadPhotoBankFundedCommoditiesColumnNo = headerRowValues.indexOf(
    UPLOAD_PHOTO_BANK_FUNDED_COMMODITIES
  );
  const collateralManagerNameColumnNo = headerRowValues.indexOf(
    COLLATERAL_MANAGER_NAME
  );
  const customerNameColumnNo = headerRowValues.indexOf(CUSTOMER_NAME);
  const pledgeBoardAvailableColumnNo = headerRowValues.indexOf(
    PLEDGE_BOARD_AVAILABLE
  );
  const pledgeBoardRemarksColumnNo =
    headerRowValues.indexOf(PLEDGE_BOARD_REMARKS);
  const uploadPhotoPledgeBoardColumnNo = headerRowValues.indexOf(
    UPLOAD_PHOTO_PLEDGE_BOARD
  );
  const stackCardAvailableColumnNo =
    headerRowValues.indexOf(STACK_CARD_AVAILABLE);
  const stackCardRemarksColumnNo = headerRowValues.indexOf(STACK_CARD_REMARKS);
  const uploadPhotoStackCardColumnNo = headerRowValues.indexOf(
    UPLOAD_PHOTO_STACK_CARD
  );
  const fireEquipmentAvailableColumnNo = headerRowValues.indexOf(
    FIRE_EQUIPMENT_AVAILABLE
  );
  const fireEquipmentRemarksColumnNo = headerRowValues.indexOf(
    FIRE_EQUIPMENT_REMARKS
  );
  const uploadPhotoFireEquipmentColumnNo = headerRowValues.indexOf(
    UPLOAD_PHOTO_FIRE_EQUIPMENT
  );
  const securityGuardAvailableColumnNo = headerRowValues.indexOf(
    SECURITY_GUARD_AVAILABLE
  );
  const securityGuardRemarksColumnNo = headerRowValues.indexOf(
    SECURITY_GUARD_REMARKS
  );
  const uploadPhotoSecurityGuardColumnNo = headerRowValues.indexOf(
    UPLOAD_PHOTO_SECURITY_GUARD
  );
  const supervisorAvailableColumnNo =
    headerRowValues.indexOf(SUPERVISOR_AVAILABLE);
  const supervisorRemarksColumnNo = headerRowValues.indexOf(SUPERVISOR_REMARKS);
  const uploadPhotoSupervisorColumnNo = headerRowValues.indexOf(
    UPLOAD_PHOTO_SUPERVISOR
  );
  const warehousePersonalIdCardAvailableColumnNo = headerRowValues.indexOf(
    WAREHOUSE_PERSONAL_ID_CARD_AVAILABLE
  );
  const warehousePersonalIdCardRemarksColumnNo = headerRowValues.indexOf(
    WAREHOUSE_PERSONAL_ID_CARD_REMARKS
  );
  const uploadPhotoWarehousePersonalIdCardColumnNo = headerRowValues.indexOf(
    UPLOAD_PHOTO_WAREHOUSE_PERSONAL_ID_CARD
  );
  const lockAndKeyInControlColumnNo = headerRowValues.indexOf(
    LOCK_AND_KEY_IN_CONTROL
  );
  const lockAndKeyRemarksColumnNo =
    headerRowValues.indexOf(LOCK_AND_KEY_REMARKS);
  const visitorsRegisterAvailableColumnNo = headerRowValues.indexOf(
    VISITORS_REGISTER_AVAILABLE
  );
  const visitorsRegisterRemarksColumnNo = headerRowValues.indexOf(
    VISITORS_REGISTER_REMARKS
  );
  const uploadPhotoVisitorsRegisterColumnNo = headerRowValues.indexOf(
    UPLOAD_PHOTO_VISITORS_REGISTER
  );
  const stockRegisterAvailableColumnNo = headerRowValues.indexOf(
    STOCK_REGISTER_AVAILABLE
  );
  const stockRegisterRemarksColumnNo = headerRowValues.indexOf(
    STOCK_REGISTER_REMARKS
  );
  const uploadPhotoStockRegisterColumnNo = headerRowValues.indexOf(
    UPLOAD_PHOTO_STOCK_REGISTER
  );
  const securityGuardAttendanceRegisterAvailableColumnNo =
    headerRowValues.indexOf(SECURITY_GUARD_ATTENDANCE_REGISTER_AVAILABLE);
  const securityGuardAttendanceRemarksColumnNo = headerRowValues.indexOf(
    SECURITY_GUARD_ATTENDANCE_REMARKS
  );
  const uploadPhotoSecurityGuardAttendanceColumnNo = headerRowValues.indexOf(
    UPLOAD_PHOTO_SECURITY_GUARD_ATTENDANCE
  );
  const supervisorAttendanceRegisterAvailableColumnNo = headerRowValues.indexOf(
    SUPERVISOR_ATTENDANCE_REGISTER_AVAILABLE
  );
  const supervisorAttendanceRemarksColumnNo = headerRowValues.indexOf(
    SUPERVISOR_ATTENDANCE_REMARKS
  );
  const uploadPhotoSupervisorAttendanceColumnNo = headerRowValues.indexOf(
    UPLOAD_PHOTO_SUPERVISOR_ATTENDANCE
  );
  const leinRegisterAvailableColumnNo = headerRowValues.indexOf(
    LEIN_REGISTER_AVAILABLE
  );
  const leinRegisterRemarksColumnNo = headerRowValues.indexOf(
    LEIN_REGISTER_REMARKS
  );
  const uploadPhotoLeinRegisterColumnNo = headerRowValues.indexOf(
    UPLOAD_PHOTO_LEIN_REGISTER
  );
  const typeOfWarehouseColumnNo = headerRowValues.indexOf(TYPE_OF_WAREHOUSE);
  const specifyWarehouseTypeColumnNo = headerRowValues.indexOf(
    SPECIFY_WAREHOUSE_TYPE
  );
  const warehouseStockBankNameColumnNo = headerRowValues.indexOf(
    WAREHOUSE_STOCK_BANK_NAME
  );
  const validLicenseAvailableColumnNo = headerRowValues.indexOf(
    VALID_LICENSE_AVAILABLE
  );
  const validLicenseRemarksColumnNo = headerRowValues.indexOf(
    VALID_LICENSE_REMARKS
  );
  const totalWarehouseCapacityColumnNo = headerRowValues.indexOf(
    TOTAL_WAREHOUSE_CAPACITY
  );
  const conditionOfRoofColumnNo = headerRowValues.indexOf(CONDITION_OF_ROOF);
  const conditionOfRoofRemarksColumnNo = headerRowValues.indexOf(
    CONDITION_OF_ROOF_REMARKS
  );
  const uploadPhotoConditionOfRoofColumnNo = headerRowValues.indexOf(
    UPLOAD_PHOTO_CONDITION_OF_ROOF
  );
  const plinthHeightColumnNo = headerRowValues.indexOf(PLINTH_HEIGHT);
  const plinthHeightRemarksColumnNo = headerRowValues.indexOf(
    PLINTH_HEIGHT_REMARKS
  );
  const uploadPhotoPlinthHeightColumnNo = headerRowValues.indexOf(
    UPLOAD_PHOTO_PLINTH_HEIGHT
  );
  const numberOfLocksColumnNo = headerRowValues.indexOf(NUMBER_OF_LOCKS);
  const numberOfGatesShuttersColumnNo = headerRowValues.indexOf(
    NUMBER_OF_GATES_SHUTTERS
  );
  const ventilationAvailableColumnNo = headerRowValues.indexOf(
    VENTILATION_AVAILABLE
  );
  const ventilationRemarksColumnNo =
    headerRowValues.indexOf(VENTILATION_REMARKS);
  const uploadPhotoVentilationColumnNo = headerRowValues.indexOf(
    UPLOAD_PHOTO_VENTILATION
  );
  const conditionOfWarehouseStructureColumnNo = headerRowValues.indexOf(
    CONDITION_OF_WAREHOUSE_STRUCTURE
  );
  const conditionOfWarehouseStructureRemarksColumnNo = headerRowValues.indexOf(
    CONDITION_OF_WAREHOUSE_STRUCTURE_REMARKS
  );
  const uploadPhotoConditionOfStructureColumnNo = headerRowValues.indexOf(
    UPLOAD_PHOTO_CONDITION_OF_STRUCTURE
  );
  const leakageInsideWarehouseColumnNo = headerRowValues.indexOf(
    LEAKAGE_INSIDE_WAREHOUSE
  );
  const leakageInsideRemarksColumnNo = headerRowValues.indexOf(
    LEAKAGE_INSIDE_REMARKS
  );
  const uploadPhotoLeakageInsideColumnNo = headerRowValues.indexOf(
    UPLOAD_PHOTO_LEAKAGE_INSIDE
  );
  const otherPartyDetailsColumnNo =
    headerRowValues.indexOf(OTHER_PARTY_DETAILS);
  const remarksAvgBagWeightColumnNo = headerRowValues.indexOf(
    REMARKS_AVG_BAG_WEIGHT
  );
  const electricalConnectionInsideWarehouseColumnNo = headerRowValues.indexOf(
    ELECTRICAL_CONNECTION_INSIDE_WAREHOUSE
  );
  const electricalConnectionSecurityColumnNo = headerRowValues.indexOf(
    ELECTRICAL_CONNECTION_SECURITY
  );
  const totalQuantityBankMTColumnNo = headerRowValues.indexOf(
    TOTAL_QUANTITY_BANK_MT
  );
  const totalBagsBankColumnNo = headerRowValues.indexOf(TOTAL_BAGS_BANK);
  const totalQuantityAgNextMTColumnNo = headerRowValues.indexOf(
    TOTAL_QUANTITY_AGNEXT_MT
  );
  const totalBagsAgNextColumnNo = headerRowValues.indexOf(TOTAL_BAGS_AGNEXT);
  const quantityMismatchColumnNo = headerRowValues.indexOf(QUANTITY_MISMATCH);
  const mismatchExcessOrLesserColumnNo = headerRowValues.indexOf(
    MISMATCH_EXCESS_OR_LESSER
  );
  const quantityMismatchQuantityMTColumnNo = headerRowValues.indexOf(
    QUANTITY_MISMATCH_QUANTITY_MT
  );
  const lesserQuantityReasonColumnNo = headerRowValues.indexOf(
    LESSER_QUANTITY_REASON
  );
  const physicalShortageNoOfBagsColumnNo = headerRowValues.indexOf(
    PHYSICAL_SHORTAGE_NO_OF_BAGS
  );
  const physicalShortageQuantityMTColumnNo = headerRowValues.indexOf(
    PHYSICAL_SHORTAGE_QUANTITY_MT
  );
  const physicalShortageRemarksColumnNo = headerRowValues.indexOf(
    PHYSICAL_SHORTAGE_REMARKS
  );
  const doCopyAvailableColumnNo = headerRowValues.indexOf(DO_COPY_AVAILABLE);
  const doCopyRemarksColumnNo = headerRowValues.indexOf(DO_COPY_REMARKS);
  const doCopyDateColumnNo = headerRowValues.indexOf(DO_COPY_DATE);
  const doCopyNumberColumnNo = headerRowValues.indexOf(DO_COPY_NUMBER);
  const uploadPhotoDoCopyColumnNo =
    headerRowValues.indexOf(UPLOAD_PHOTO_DO_COPY);
  const whrNumberColumnNo = headerRowValues.indexOf(WHR_NUMBER);
  const whrDateColumnNo = headerRowValues.indexOf(WHR_DATE);
  const whrDetailsNoOfBagsColumnNo = headerRowValues.indexOf(
    WHR_DETAILS_NO_OF_BAGS
  );
  const whrDetailsQuantityMTColumnNo = headerRowValues.indexOf(
    WHR_DETAILS_QUANTITY_MT
  );
  const whrDetailsPhotoColumnNo = headerRowValues.indexOf(WHR_DETAILS_PHOTO);
  const discrepancyColumnNo = headerRowValues.indexOf(DISCREPANCY);
  const inspectionReportColumnNo = headerRowValues.indexOf(
    INSPECTION_REPORT_COLUMN_NAME
  );
  const pdfLinkColumnNo = headerRowValues.indexOf(PDF_LINK_COLUMN_NAME);
  const triggerColumnNo = headerRowValues.indexOf(TRIGGER);
  let timestamp = sheetValues[rowNumber][timestampColumnNo];

  const dateObj = new Date(timestamp);
  const date = dateObj.toLocaleDateString();
  const time = dateObj.toLocaleTimeString();
  const timeForObj = date + " " + time;

  formattedDate = Utilities.formatDate(new Date(timestamp), timezone, format);
  timestamp = formattedDate;

  let auditDate = sheetValues[rowNumber][auditDateColumnNo];
  formattedDate = Utilities.formatDate(new Date(auditDate), timezone, format);
  auditDate = formattedDate;

  const auditorName = sheetValues[rowNumber][auditorNameColumnNo];
  const warehouseCode = sheetValues[rowNumber][warehouseCodeColumnNo];
  const commodityHealth = sheetValues[rowNumber][commodityHealthColumnNo];
  const commoditiesAvailable =
    sheetValues[rowNumber][commoditiesAvailableColumnNo];
  const uploadPhotoCommodity =
    sheetValues[rowNumber][uploadPhotoCommodityColumnNo];
  const infestationDeteriorationNoticed =
    sheetValues[rowNumber][infestationDeteriorationNoticedColumnNo];
  const infestationDeteriorationRemarks =
    sheetValues[rowNumber][infestationDeteriorationRemarksColumnNo];
  const uploadPhotoInfestationDeterioration =
    sheetValues[rowNumber][uploadPhotoInfestationDeteriorationColumnNo];
  const fumigationRequired = sheetValues[rowNumber][fumigationRequiredColumnNo];
  const fumigationRequiredRemarks =
    sheetValues[rowNumber][fumigationRequiredRemarksColumnNo];
  const uploadPhotoFumigation =
    sheetValues[rowNumber][uploadPhotoFumigationColumnNo];
  const dunnageAvailable = sheetValues[rowNumber][dunnageAvailableColumnNo];
  const dunnageAvailableRemarks =
    sheetValues[rowNumber][dunnageAvailableRemarksColumnNo];
  const uploadPhotoDunnage = sheetValues[rowNumber][uploadPhotoDunnageColumnNo];
  const stockKeptCountable = sheetValues[rowNumber][stockKeptCountableColumnNo];
  const stockKeptCountableRemarks =
    sheetValues[rowNumber][stockKeptCountableRemarksColumnNo];
  const uploadPhotoStockKeptCountable =
    sheetValues[rowNumber][uploadPhotoStockKeptCountableColumnNo];
  const weighingScaleAvailable =
    sheetValues[rowNumber][weighingScaleAvailableColumnNo];
  const weighingScaleAvailableRemarks =
    sheetValues[rowNumber][weighingScaleAvailableRemarksColumnNo];
  const uploadPhotoWeighingScale =
    sheetValues[rowNumber][uploadPhotoWeighingScaleColumnNo];
  const hygieneCleanlinessMaintained =
    sheetValues[rowNumber][hygieneCleanlinessMaintainedColumnNo];
  const hygieneCleanlinessMaintainedRemarks =
    sheetValues[rowNumber][hygieneCleanlinessMaintainedRemarksColumnNo];
  const uploadPhotoHygieneCleanliness =
    sheetValues[rowNumber][uploadPhotoHygieneCleanlinessColumnNo];
  const bankFundedCommoditiesStored =
    sheetValues[rowNumber][bankFundedCommoditiesStoredColumnNo];
  const bankFundedCommoditiesDetails =
    sheetValues[rowNumber][bankFundedCommoditiesDetailsColumnNo];
  const uploadPhotoBankFundedCommodities =
    sheetValues[rowNumber][uploadPhotoBankFundedCommoditiesColumnNo];
  const collateralManagerName =
    sheetValues[rowNumber][collateralManagerNameColumnNo];
  const customerName = sheetValues[rowNumber][customerNameColumnNo];
  const pledgeBoardAvailable =
    sheetValues[rowNumber][pledgeBoardAvailableColumnNo];
  const pledgeBoardRemarks = sheetValues[rowNumber][pledgeBoardRemarksColumnNo];
  const uploadPhotoPledgeBoard =
    sheetValues[rowNumber][uploadPhotoPledgeBoardColumnNo];
  const stackCardAvailable = sheetValues[rowNumber][stackCardAvailableColumnNo];
  const stackCardRemarks = sheetValues[rowNumber][stackCardRemarksColumnNo];
  const uploadPhotoStackCard =
    sheetValues[rowNumber][uploadPhotoStackCardColumnNo];
  const fireEquipmentAvailable =
    sheetValues[rowNumber][fireEquipmentAvailableColumnNo];
  const fireEquipmentRemarks =
    sheetValues[rowNumber][fireEquipmentRemarksColumnNo];
  const uploadPhotoFireEquipment =
    sheetValues[rowNumber][uploadPhotoFireEquipmentColumnNo];
  const securityGuardAvailable =
    sheetValues[rowNumber][securityGuardAvailableColumnNo];
  const securityGuardRemarks =
    sheetValues[rowNumber][securityGuardRemarksColumnNo];
  const uploadPhotoSecurityGuard =
    sheetValues[rowNumber][uploadPhotoSecurityGuardColumnNo];
  const supervisorAvailable =
    sheetValues[rowNumber][supervisorAvailableColumnNo];
  const supervisorRemarks = sheetValues[rowNumber][supervisorRemarksColumnNo];
  const uploadPhotoSupervisor =
    sheetValues[rowNumber][uploadPhotoSupervisorColumnNo];
  const warehousePersonalIdCardAvailable =
    sheetValues[rowNumber][warehousePersonalIdCardAvailableColumnNo];
  const warehousePersonalIdCardRemarks =
    sheetValues[rowNumber][warehousePersonalIdCardRemarksColumnNo];
  const uploadPhotoWarehousePersonalIdCard =
    sheetValues[rowNumber][uploadPhotoWarehousePersonalIdCardColumnNo];
  const lockAndKeyInControl =
    sheetValues[rowNumber][lockAndKeyInControlColumnNo];
  const lockAndKeyRemarks = sheetValues[rowNumber][lockAndKeyRemarksColumnNo];
  const visitorsRegisterAvailable =
    sheetValues[rowNumber][visitorsRegisterAvailableColumnNo];
  const visitorsRegisterRemarks =
    sheetValues[rowNumber][visitorsRegisterRemarksColumnNo];
  const uploadPhotoVisitorsRegister =
    sheetValues[rowNumber][uploadPhotoVisitorsRegisterColumnNo];
  const stockRegisterAvailable =
    sheetValues[rowNumber][stockRegisterAvailableColumnNo];
  const stockRegisterRemarks =
    sheetValues[rowNumber][stockRegisterRemarksColumnNo];
  const uploadPhotoStockRegister =
    sheetValues[rowNumber][uploadPhotoStockRegisterColumnNo];
  const securityGuardAttendanceRegisterAvailable =
    sheetValues[rowNumber][securityGuardAttendanceRegisterAvailableColumnNo];
  const securityGuardAttendanceRemarks =
    sheetValues[rowNumber][securityGuardAttendanceRemarksColumnNo];
  const uploadPhotoSecurityGuardAttendance =
    sheetValues[rowNumber][uploadPhotoSecurityGuardAttendanceColumnNo];
  const supervisorAttendanceRegisterAvailable =
    sheetValues[rowNumber][supervisorAttendanceRegisterAvailableColumnNo];
  const supervisorAttendanceRemarks =
    sheetValues[rowNumber][supervisorAttendanceRemarksColumnNo];
  const uploadPhotoSupervisorAttendance =
    sheetValues[rowNumber][uploadPhotoSupervisorAttendanceColumnNo];
  const leinRegisterAvailable =
    sheetValues[rowNumber][leinRegisterAvailableColumnNo];
  const leinRegisterRemarks =
    sheetValues[rowNumber][leinRegisterRemarksColumnNo];
  const uploadPhotoLeinRegister =
    sheetValues[rowNumber][uploadPhotoLeinRegisterColumnNo];
  const typeOfWarehouse = sheetValues[rowNumber][typeOfWarehouseColumnNo];
  const specifyWarehouseType =
    sheetValues[rowNumber][specifyWarehouseTypeColumnNo];
  const warehouseStockBankName =
    sheetValues[rowNumber][warehouseStockBankNameColumnNo];
  const validLicenseAvailable =
    sheetValues[rowNumber][validLicenseAvailableColumnNo];
  const validLicenseRemarks =
    sheetValues[rowNumber][validLicenseRemarksColumnNo];
  const totalWarehouseCapacity =
    sheetValues[rowNumber][totalWarehouseCapacityColumnNo];
  const conditionOfRoof = sheetValues[rowNumber][conditionOfRoofColumnNo];
  const conditionOfRoofRemarks =
    sheetValues[rowNumber][conditionOfRoofRemarksColumnNo];
  const uploadPhotoConditionOfRoof =
    sheetValues[rowNumber][uploadPhotoConditionOfRoofColumnNo];
  const plinthHeight = sheetValues[rowNumber][plinthHeightColumnNo];
  const plinthHeightRemarks =
    sheetValues[rowNumber][plinthHeightRemarksColumnNo];
  const uploadPhotoPlinthHeight =
    sheetValues[rowNumber][uploadPhotoPlinthHeightColumnNo];
  const numberOfLocks = sheetValues[rowNumber][numberOfLocksColumnNo];
  const numberOfGatesShutters =
    sheetValues[rowNumber][numberOfGatesShuttersColumnNo];
  const ventilationAvailable =
    sheetValues[rowNumber][ventilationAvailableColumnNo];
  const ventilationRemarks = sheetValues[rowNumber][ventilationRemarksColumnNo];
  const uploadPhotoVentilation =
    sheetValues[rowNumber][uploadPhotoVentilationColumnNo];
  const conditionOfWarehouseStructure =
    sheetValues[rowNumber][conditionOfWarehouseStructureColumnNo];
  const conditionOfWarehouseStructureRemarks =
    sheetValues[rowNumber][conditionOfWarehouseStructureRemarksColumnNo];
  const uploadPhotoConditionOfStructure =
    sheetValues[rowNumber][uploadPhotoConditionOfStructureColumnNo];
  const leakageInsideWarehouse =
    sheetValues[rowNumber][leakageInsideWarehouseColumnNo];
  const leakageInsideRemarks =
    sheetValues[rowNumber][leakageInsideRemarksColumnNo];
  const uploadPhotoLeakageInside =
    sheetValues[rowNumber][uploadPhotoLeakageInsideColumnNo];
  const otherPartyDetails = sheetValues[rowNumber][otherPartyDetailsColumnNo];
  const remarksAvgBagWeight =
    sheetValues[rowNumber][remarksAvgBagWeightColumnNo];
  const electricalConnectionInsideWarehouse =
    sheetValues[rowNumber][electricalConnectionInsideWarehouseColumnNo];
  const electricalConnectionSecurity =
    sheetValues[rowNumber][electricalConnectionSecurityColumnNo];
  const totalQuantityBankMT =
    sheetValues[rowNumber][totalQuantityBankMTColumnNo];
  const totalBagsBank = sheetValues[rowNumber][totalBagsBankColumnNo];
  const totalQuantityAgNextMT =
    sheetValues[rowNumber][totalQuantityAgNextMTColumnNo];
  const totalBagsAgNext = sheetValues[rowNumber][totalBagsAgNextColumnNo];
  const quantityMismatch = sheetValues[rowNumber][quantityMismatchColumnNo];
  const mismatchExcessOrLesser =
    sheetValues[rowNumber][mismatchExcessOrLesserColumnNo];
  const quantityMismatchQuantityMT =
    sheetValues[rowNumber][quantityMismatchQuantityMTColumnNo];
  const lesserQuantityReason =
    sheetValues[rowNumber][lesserQuantityReasonColumnNo];
  const physicalShortageNoOfBags =
    sheetValues[rowNumber][physicalShortageNoOfBagsColumnNo];
  const physicalShortageQuantityMT =
    sheetValues[rowNumber][physicalShortageQuantityMTColumnNo];
  const physicalShortageRemarks =
    sheetValues[rowNumber][physicalShortageRemarksColumnNo];
  let doCopyAvailable = sheetValues[rowNumber][doCopyAvailableColumnNo];
  if (
    typeof doCopyAvailable == "undefined" ||
    doCopyAvailable.toString() == ""
  ) {
    doCopyAvailable = "-";
  }

  const doCopyRemarks = sheetValues[rowNumber][doCopyRemarksColumnNo];
  let doCopyDate = sheetValues[rowNumber][doCopyDateColumnNo];
  if (typeof doCopyDate == "undefined" || doCopyDate.toString() == "") {
    doCopyDate = "-";
  } else if (typeof doCopyDate == "object") {
    formattedDate = Utilities.formatDate(
      new Date(doCopyDate),
      timezone,
      format
    );
    doCopyDate = formattedDate;
  }

  let doCopyNumber = sheetValues[rowNumber][doCopyNumberColumnNo];
  if (typeof doCopyNumber == "undefined" || doCopyNumber.toString() == "") {
    doCopyNumber = "-";
  }

  const uploadPhotoDoCopy = sheetValues[rowNumber][uploadPhotoDoCopyColumnNo];
  let whrNumber = sheetValues[rowNumber][whrNumberColumnNo];
  if (typeof whrNumber == "undefined" || whrNumber.toString() == "") {
    whrNumber = "-";
  }

  let whrDate = sheetValues[rowNumber][whrDateColumnNo];
  if (typeof whrDate == "undefined" || whrDate.toString() == "") {
    whrDate = "-";
  } else if (typeof whrDate == "object") {
    formattedDate = Utilities.formatDate(new Date(whrDate), timezone, format);
    whrDate = formattedDate;
  }

  let whrDetailsNoOfBags = sheetValues[rowNumber][whrDetailsNoOfBagsColumnNo];
  if (
    typeof whrDetailsNoOfBags == "undefined" ||
    whrDetailsNoOfBags.toString() == ""
  ) {
    whrDetailsNoOfBags = "-";
  }

  let whrDetailsQuantityMT =
    sheetValues[rowNumber][whrDetailsQuantityMTColumnNo];
  if (
    typeof whrDetailsQuantityMT == "undefined" ||
    whrDetailsQuantityMT.toString() == ""
  ) {
    whrDetailsQuantityMT = "-";
  }

  let whrDetailsPhoto = sheetValues[rowNumber][whrDetailsPhotoColumnNo];
  if (
    typeof whrDetailsPhoto == "undefined" ||
    whrDetailsPhoto.toString() == ""
  ) {
    whrDetailsPhoto = "-";
  }

  const discrepancy = sheetValues[rowNumber][discrepancyColumnNo];
  const uploadPhotoinspectionReport =
    sheetValues[rowNumber][inspectionReportColumnNo];
  const pdfLink = sheetValues[rowNumber][pdfLinkColumnNo];
  const trigger = sheetValues[rowNumber][triggerColumnNo];
  console.log(trigger);
  let values = [
    timeForObj,
    auditDate,
    auditorName,
    warehouseCode,
    commodityHealth,
    commoditiesAvailable,
    uploadPhotoCommodity,
    infestationDeteriorationNoticed,
    infestationDeteriorationRemarks,
    uploadPhotoInfestationDeterioration,
    fumigationRequired,
    fumigationRequiredRemarks,
    uploadPhotoFumigation,
    dunnageAvailable,
    dunnageAvailableRemarks,
    uploadPhotoDunnage,
    stockKeptCountable,
    stockKeptCountableRemarks,
    uploadPhotoStockKeptCountable,
    weighingScaleAvailable,
    weighingScaleAvailableRemarks,
    uploadPhotoWeighingScale,
    hygieneCleanlinessMaintained,
    hygieneCleanlinessMaintainedRemarks,
    uploadPhotoHygieneCleanliness,
    bankFundedCommoditiesStored,
    bankFundedCommoditiesDetails,
    uploadPhotoBankFundedCommodities,
    collateralManagerName,
    customerName,
    pledgeBoardAvailable,
    pledgeBoardRemarks,
    uploadPhotoPledgeBoard,
    stackCardAvailable,
    stackCardRemarks,
    uploadPhotoStackCard,
    fireEquipmentAvailable,
    fireEquipmentRemarks,
    uploadPhotoFireEquipment,
    securityGuardAvailable,
    securityGuardRemarks,
    uploadPhotoSecurityGuard,
    supervisorAvailable,
    supervisorRemarks,
    uploadPhotoSupervisor,
    warehousePersonalIdCardAvailable,
    warehousePersonalIdCardRemarks,
    uploadPhotoWarehousePersonalIdCard,
    lockAndKeyInControl,
    lockAndKeyRemarks,
    visitorsRegisterAvailable,
    visitorsRegisterRemarks,
    uploadPhotoVisitorsRegister,
    stockRegisterAvailable,
    stockRegisterRemarks,
    uploadPhotoStockRegister,
    securityGuardAttendanceRegisterAvailable,
    securityGuardAttendanceRemarks,
    uploadPhotoSecurityGuardAttendance,
    supervisorAttendanceRegisterAvailable,
    supervisorAttendanceRemarks,
    uploadPhotoSupervisorAttendance,
    leinRegisterAvailable,
    leinRegisterRemarks,
    uploadPhotoLeinRegister,
    typeOfWarehouse,
    specifyWarehouseType,
    warehouseStockBankName,
    validLicenseAvailable,
    validLicenseRemarks,
    totalWarehouseCapacity,
    conditionOfRoof,
    conditionOfRoofRemarks,
    uploadPhotoConditionOfRoof,
    plinthHeight,
    plinthHeightRemarks,
    uploadPhotoPlinthHeight,
    numberOfLocks,
    numberOfGatesShutters,
    ventilationAvailable,
    ventilationRemarks,
    uploadPhotoVentilation,
    conditionOfWarehouseStructure,
    conditionOfWarehouseStructureRemarks,
    uploadPhotoConditionOfStructure,
    leakageInsideWarehouse,
    leakageInsideRemarks,
    uploadPhotoLeakageInside,
    otherPartyDetails,
    remarksAvgBagWeight,
    electricalConnectionInsideWarehouse,
    electricalConnectionSecurity,
    totalQuantityBankMT,
    totalBagsBank,
    totalQuantityAgNextMT,
    totalBagsAgNext,
    quantityMismatch,
    mismatchExcessOrLesser,
    quantityMismatchQuantityMT,
    lesserQuantityReason,
    physicalShortageNoOfBags,
    physicalShortageQuantityMT,
    physicalShortageRemarks,
    doCopyAvailable,
    doCopyRemarks,
    doCopyDate,
    doCopyNumber,
    uploadPhotoDoCopy,
    whrNumber,
    whrDate,
    whrDetailsNoOfBags,
    whrDetailsQuantityMT,
    whrDetailsPhoto,
    discrepancy,
    uploadPhotoinspectionReport,
    pdfLink,
  ];

  const requiredArrOfObject = await objectGenerator(values);
  // console.log(requiredArrOfObject)
  await objectToAzureEndPoint(requiredArrOfObject, rowNumber);

  body.replaceText("<<" + TIMESTAMP + ">>", timestamp + " ");
  body.replaceText("<<" + AUDIT_DATE + ">>", auditDate + " ");
  body.replaceText("<<" + AUDITOR_NAME + ">>", auditorName + " ");
  body.replaceText("<<" + WAREHOUSE_CODE + ">>", warehouseCode + " ");
  body.replaceText("<<" + COMMODITY_HEALTH + ">>", commodityHealth + " ");
  body.replaceText(
    "<<" + COMMODITIES_AVAILABLE + ">>",
    commoditiesAvailable + " "
  );
  body.replaceText(
    "<<" + INFESTATION_DETERIORATION_NOTICED + ">>",
    infestationDeteriorationNoticed + " "
  );
  body.replaceText(
    "<<" + INFESTATION_DETERIORATION_REMARKS + ">>",
    infestationDeteriorationRemarks + " "
  );
  body.replaceText("<<" + FUMIGATION_REQUIRED + ">>", fumigationRequired + " ");
  body.replaceText(
    "<<" + FUMIGATION_REQUIRED_REMARKS + ">>",
    fumigationRequiredRemarks + " "
  );
  body.replaceText("<<" + DUNNAGE_AVAILABLE + ">>", dunnageAvailable + " ");
  body.replaceText(
    "<<" + DUNNAGE_AVAILABLE_REMARKS + ">>",
    dunnageAvailableRemarks + " "
  );
  body.replaceText(
    "<<" + STOCK_KEPT_COUNTABLE + ">>",
    stockKeptCountable + " "
  );
  body.replaceText(
    "<<" + STOCK_KEPT_COUNTABLE_REMARKS + ">>",
    stockKeptCountableRemarks + " "
  );
  body.replaceText(
    "<<" + WEIGHING_SCALE_AVAILABLE + ">>",
    weighingScaleAvailable + " "
  );
  body.replaceText(
    "<<" + WEIGHING_SCALE_AVAILABLE_REMARKS + ">>",
    weighingScaleAvailableRemarks + " "
  );
  body.replaceText(
    "<<" + HYGIENE_CLEANLINESS_MAINTAINED + ">>",
    hygieneCleanlinessMaintained + " "
  );
  body.replaceText(
    "<<" + HYGIENE_CLEANLINESS_MAINTAINED_REMARKS + ">>",
    hygieneCleanlinessMaintainedRemarks + " "
  );
  body.replaceText(
    "<<" + BANK_FUNDED_COMMODITIES_STORED + ">>",
    bankFundedCommoditiesStored + " "
  );
  body.replaceText(
    "<<" + BANK_FUNDED_COMMODITIES_DETAILS + ">>",
    bankFundedCommoditiesDetails + " "
  );
  body.replaceText(
    "<<" + COLLATERAL_MANAGER_NAME + ">>",
    collateralManagerName + " "
  );
  body.replaceText("<<" + CUSTOMER_NAME + ">>", customerName + " ");
  body.replaceText(
    "<<" + PLEDGE_BOARD_AVAILABLE + ">>",
    pledgeBoardAvailable + " "
  );
  body.replaceText(
    "<<" + PLEDGE_BOARD_REMARKS + ">>",
    pledgeBoardRemarks + " "
  );
  body.replaceText(
    "<<" + STACK_CARD_AVAILABLE + ">>",
    stackCardAvailable + " "
  );
  body.replaceText("<<" + STACK_CARD_REMARKS + ">>", stackCardRemarks + " ");
  body.replaceText(
    "<<" + FIRE_EQUIPMENT_AVAILABLE + ">>",
    fireEquipmentAvailable + " "
  );
  body.replaceText(
    "<<" + FIRE_EQUIPMENT_REMARKS + ">>",
    fireEquipmentRemarks + " "
  );
  body.replaceText(
    "<<" + SECURITY_GUARD_AVAILABLE + ">>",
    securityGuardAvailable + " "
  );
  body.replaceText(
    "<<" + SECURITY_GUARD_REMARKS + ">>",
    securityGuardRemarks + " "
  );
  body.replaceText(
    "<<" + SUPERVISOR_AVAILABLE + ">>",
    supervisorAvailable + " "
  );
  body.replaceText("<<" + SUPERVISOR_REMARKS + ">>", supervisorRemarks + " ");
  body.replaceText(
    "<<" + WAREHOUSE_PERSONAL_ID_CARD_AVAILABLE + ">>",
    warehousePersonalIdCardAvailable + " "
  );
  body.replaceText(
    "<<" + WAREHOUSE_PERSONAL_ID_CARD_REMARKS + ">>",
    warehousePersonalIdCardRemarks + " "
  );
  body.replaceText(
    "<<" + LOCK_AND_KEY_IN_CONTROL + ">>",
    lockAndKeyInControl + " "
  );
  body.replaceText("<<" + LOCK_AND_KEY_REMARKS + ">>", lockAndKeyRemarks + " ");
  body.replaceText(
    "<<" + VISITORS_REGISTER_AVAILABLE + ">>",
    visitorsRegisterAvailable + " "
  );
  body.replaceText(
    "<<" + VISITORS_REGISTER_REMARKS + ">>",
    visitorsRegisterRemarks + " "
  );
  body.replaceText(
    "<<" + STOCK_REGISTER_AVAILABLE + ">>",
    stockRegisterAvailable + " "
  );
  body.replaceText(
    "<<" + STOCK_REGISTER_REMARKS + ">>",
    stockRegisterRemarks + " "
  );
  body.replaceText(
    "<<" + SECURITY_GUARD_ATTENDANCE_REGISTER_AVAILABLE + ">>",
    securityGuardAttendanceRegisterAvailable + " "
  );
  body.replaceText(
    "<<" + SECURITY_GUARD_ATTENDANCE_REMARKS + ">>",
    securityGuardAttendanceRemarks + " "
  );
  body.replaceText(
    "<<" + SUPERVISOR_ATTENDANCE_REGISTER_AVAILABLE + ">>",
    supervisorAttendanceRegisterAvailable + " "
  );
  body.replaceText(
    "<<" + SUPERVISOR_ATTENDANCE_REMARKS + ">>",
    supervisorAttendanceRemarks + " "
  );
  body.replaceText(
    "<<" + LEIN_REGISTER_AVAILABLE + ">>",
    leinRegisterAvailable + " "
  );
  body.replaceText(
    "<<" + LEIN_REGISTER_REMARKS + ">>",
    leinRegisterRemarks + " "
  );
  body.replaceText("<<" + TYPE_OF_WAREHOUSE + ">>", typeOfWarehouse + " ");
  body.replaceText(
    "<<" + SPECIFY_WAREHOUSE_TYPE + ">>",
    specifyWarehouseType + " "
  );
  body.replaceText(
    "<<" + WAREHOUSE_STOCK_BANK_NAME + ">>",
    warehouseStockBankName + " "
  );
  body.replaceText(
    "<<" + VALID_LICENSE_AVAILABLE + ">>",
    validLicenseAvailable + " "
  );
  body.replaceText(
    "<<" + VALID_LICENSE_REMARKS + ">>",
    validLicenseRemarks + " "
  );
  body.replaceText(
    "<<" + TOTAL_WAREHOUSE_CAPACITY + ">>",
    totalWarehouseCapacity + " "
  );
  body.replaceText("<<" + CONDITION_OF_ROOF + ">>", conditionOfRoof + " ");
  body.replaceText(
    "<<" + CONDITION_OF_ROOF_REMARKS + ">>",
    conditionOfRoofRemarks + " "
  );
  body.replaceText("<<" + PLINTH_HEIGHT + ">>", plinthHeight + " ");
  body.replaceText(
    "<<" + PLINTH_HEIGHT_REMARKS + ">>",
    plinthHeightRemarks + " "
  );
  body.replaceText("<<" + NUMBER_OF_LOCKS + ">>", numberOfLocks + " ");
  body.replaceText(
    "<<" + NUMBER_OF_GATES_SHUTTERS + ">>",
    numberOfGatesShutters + " "
  );
  body.replaceText(
    "<<" + VENTILATION_AVAILABLE + ">>",
    ventilationAvailable + " "
  );
  body.replaceText("<<" + VENTILATION_REMARKS + ">>", ventilationRemarks + " ");
  body.replaceText(
    "<<" + CONDITION_OF_WAREHOUSE_STRUCTURE + ">>",
    conditionOfWarehouseStructure + " "
  );
  body.replaceText(
    "<<" + CONDITION_OF_WAREHOUSE_STRUCTURE_REMARKS + ">>",
    conditionOfWarehouseStructureRemarks + " "
  );
  body.replaceText(
    "<<" + LEAKAGE_INSIDE_WAREHOUSE + ">>",
    leakageInsideWarehouse + " "
  );
  body.replaceText(
    "<<" + LEAKAGE_INSIDE_REMARKS + ">>",
    leakageInsideRemarks + " "
  );
  body.replaceText("<<" + OTHER_PARTY_DETAILS + ">>", otherPartyDetails + " ");
  body.replaceText(
    "<<" + REMARKS_AVG_BAG_WEIGHT + ">>",
    remarksAvgBagWeight + " "
  );
  body.replaceText(
    "<<" + ELECTRICAL_CONNECTION_INSIDE_WAREHOUSE + ">>",
    electricalConnectionInsideWarehouse + " "
  );
  body.replaceText(
    "<<" + ELECTRICAL_CONNECTION_SECURITY + ">>",
    electricalConnectionSecurity + " "
  );
  body.replaceText(
    "<<" + TOTAL_QUANTITY_BANK_MT + ">>",
    totalQuantityBankMT + " "
  );
  body.replaceText("<<" + TOTAL_BAGS_BANK + ">>", totalBagsBank + " ");
  body.replaceText(
    "<<" + TOTAL_QUANTITY_AGNEXT_MT + ">>",
    totalQuantityAgNextMT + " "
  );
  body.replaceText("<<" + TOTAL_BAGS_AGNEXT + ">>", totalBagsAgNext + " ");
  body.replaceText("<<" + QUANTITY_MISMATCH + ">>", quantityMismatch + " ");
  body.replaceText(
    "<<" + MISMATCH_EXCESS_OR_LESSER + ">>",
    mismatchExcessOrLesser + " "
  );
  body.replaceText(
    "<<" + QUANTITY_MISMATCH_QUANTITY_MT + ">>",
    quantityMismatchQuantityMT + " "
  );
  body.replaceText(
    "<<" + LESSER_QUANTITY_REASON + ">>",
    lesserQuantityReason + " "
  );
  body.replaceText(
    "<<" + PHYSICAL_SHORTAGE_NO_OF_BAGS + ">>",
    physicalShortageNoOfBags + " "
  );
  body.replaceText(
    "<<" + PHYSICAL_SHORTAGE_QUANTITY_MT + ">>",
    physicalShortageQuantityMT + " "
  );
  body.replaceText(
    "<<" + PHYSICAL_SHORTAGE_REMARKS + ">>",
    physicalShortageRemarks + " "
  );
  body.replaceText("<<" + DO_COPY_AVAILABLE + ">>", doCopyAvailable + " ");
  body.replaceText("<<" + DO_COPY_REMARKS + ">>", doCopyRemarks + " ");
  body.replaceText("<<" + DO_COPY_DATE + ">>", doCopyDate + " ");
  body.replaceText("<<" + DO_COPY_NUMBER + ">>", doCopyNumber + " ");
  body.replaceText("<<" + WHR_NUMBER + ">>", whrNumber + " ");
  body.replaceText("<<" + WHR_DATE + ">>", whrDate + " ");
  body.replaceText(
    "<<" + WHR_DETAILS_NO_OF_BAGS + ">>",
    whrDetailsNoOfBags + " "
  );
  body.replaceText(
    "<<" + WHR_DETAILS_QUANTITY_MT + ">>",
    whrDetailsQuantityMT + " "
  );
  body.replaceText("<<" + DISCREPANCY + ">>", discrepancy + " ");

  body.replaceText("undefined", "");
  body.replaceText("null", "");

  // body.replaceText("<<" + UPLOAD_PHOTO_COMMODITY + ">>", uploadPhotoCommodity + " ");
  // body.replaceText("<<" + UPLOAD_PHOTO_INFESTATION_DETERIORATION + ">>", uploadPhotoInfestationDeterioration + " ");
  // body.replaceText("<<" + UPLOAD_PHOTO_FUMIGATION + ">>", uploadPhotoFumigation + " ");
  // body.replaceText("<<" + UPLOAD_PHOTO_DUNNAGE + ">>", uploadPhotoDunnage + " ");
  // body.replaceText("<<" + UPLOAD_PHOTO_STOCK_KEPT_COUNTABLE + ">>", uploadPhotoStockKeptCountable + " ");
  // body.replaceText("<<" + UPLOAD_PHOTO_WEIGHING_SCALE + ">>", uploadPhotoWeighingScale + " ");
  // body.replaceText("<<" + UPLOAD_PHOTO_HYGIENE_CLEANLINESS + ">>", uploadPhotoHygieneCleanliness + " ");
  // body.replaceText("<<" + UPLOAD_PHOTO_BANK_FUNDED_COMMODITIES + ">>", uploadPhotoBankFundedCommodities + " ");
  // body.replaceText("<<" + UPLOAD_PHOTO_PLEDGE_BOARD + ">>", uploadPhotoPledgeBoard + " ");
  // body.replaceText("<<" + UPLOAD_PHOTO_STACK_CARD + ">>", uploadPhotoStackCard + " ");
  // body.replaceText("<<" + UPLOAD_PHOTO_FIRE_EQUIPMENT + ">>", uploadPhotoFireEquipment + " ");
  // body.replaceText("<<" + UPLOAD_PHOTO_SECURITY_GUARD + ">>", uploadPhotoSecurityGuard + " ");
  // body.replaceText("<<" + UPLOAD_PHOTO_SUPERVISOR + ">>", uploadPhotoSupervisor + " ");
  // body.replaceText("<<" + UPLOAD_PHOTO_WAREHOUSE_PERSONAL_ID_CARD + ">>", uploadPhotoWarehousePersonalIdCard + " ");
  // body.replaceText("<<" + UPLOAD_PHOTO_VISITORS_REGISTER + ">>", uploadPhotoVisitorsRegister + " ");
  // body.replaceText("<<" + UPLOAD_PHOTO_STOCK_REGISTER + ">>", uploadPhotoStockRegister + " ");
  // body.replaceText("<<" + UPLOAD_PHOTO_SECURITY_GUARD_ATTENDANCE + ">>", uploadPhotoSecurityGuardAttendance + " ");
  // body.replaceText("<<" + UPLOAD_PHOTO_SUPERVISOR_ATTENDANCE + ">>", uploadPhotoSupervisorAttendance + " ");
  // body.replaceText("<<" + UPLOAD_PHOTO_LEIN_REGISTER + ">>", uploadPhotoLeinRegister + " ");
  // body.replaceText("<<" + UPLOAD_PHOTO_CONDITION_OF_ROOF + ">>", uploadPhotoConditionOfRoof + " ");
  // body.replaceText("<<" + UPLOAD_PHOTO_PLINTH_HEIGHT + ">>", uploadPhotoPlinthHeight + " ");
  // body.replaceText("<<" + UPLOAD_PHOTO_VENTILATION + ">>", uploadPhotoVentilation + " ");
  // body.replaceText("<<" + UPLOAD_PHOTO_CONDITION_OF_STRUCTURE + ">>", uploadPhotoConditionOfStructure + " ");
  // body.replaceText("<<" + UPLOAD_PHOTO_LEAKAGE_INSIDE + ">>", uploadPhotoLeakageInside + " ");
  // body.replaceText("<<" + UPLOAD_PHOTO_DO_COPY + ">>", uploadPhotoDoCopy + " ");
  // body.replaceText("<<" + WHR_DETAILS_PHOTO + ">>", whrDetailsPhoto + " ");

  const photosColumnsArray = {
    "Upload Photo - Commodity": uploadPhotoCommodity,
    "Upload Photo - Infestation / Deterioration":
      uploadPhotoInfestationDeterioration,
    "Upload Photo - Fumigation": uploadPhotoFumigation,
    "Upload Photo - Dunnage Available": uploadPhotoDunnage,
    "Upload Photo - Stock kept in countable position":
      uploadPhotoStockKeptCountable,
    "Upload Photo - Weighing scale": uploadPhotoWeighingScale,
    "Upload photo - Hygiene / Cleanliness maintained in warehouse":
      uploadPhotoHygieneCleanliness,
    "Upload Photo - Bank Funded commodities": uploadPhotoBankFundedCommodities,
    "Upload Photo - Pledge Board Available": uploadPhotoPledgeBoard,
    "Upload Photo - Stack card available": uploadPhotoStackCard,
    "Upload Photo - Fire equipment available": uploadPhotoFireEquipment,
    "Upload Photo - Security Guard available during visit":
      uploadPhotoSecurityGuard,
    "Upload Photo - Supervisor available during visit": uploadPhotoSupervisor,
    "Upload Photo - Id card of warehouse personal available":
      uploadPhotoWarehousePersonalIdCard,
    "Upload Photo - Visitors register available": uploadPhotoVisitorsRegister,
    "Upload Photo - Stock register available": uploadPhotoStockRegister,
    "Upload Photo - Security guard's attendance register":
      uploadPhotoSecurityGuardAttendance,
    "Upload Photo - Supervisor's attendance register":
      uploadPhotoSupervisorAttendance,
    "Upload Photo - Lein register available & Lein marked":
      uploadPhotoLeinRegister,
    "Upload Photo - Condition of Roof": uploadPhotoConditionOfRoof,
    "Upload Photo - Plinth Height 3 Feet / More": uploadPhotoPlinthHeight,
    "Upload Photo - Ventilation Available": uploadPhotoVentilation,
    "Upload Photo - Condition of warehouse structure":
      uploadPhotoConditionOfStructure,
    "Upload Photo - Leakage inside the Warehouse": uploadPhotoLeakageInside,
    "Upload photo - DO Copy": uploadPhotoDoCopy,
    "WHR Details - Photo": whrDetailsPhoto,
  };
  const tables = doc.getTables();
  for (let i = 0; i < tables.length - 1; i++) {
    const table = tables[i];
    const tableRows = table.getNumRows();

    for (let j = 0; j < tableRows; j++) {
      for (let k = 0; k < table.getRow(j).getNumCells(); k++) {
        const cell = table.getCell(j, k);
        for (let columnName in photosColumnsArray) {
          if (photosColumnsArray[columnName] == "-") {
            body.replaceText("<<" + columnName + ">>", "-");
            continue;
          }
          const imageAdded = await replaceTextWithImagesInCell(
            cell,
            body,
            columnName,
            photosColumnsArray[columnName]
          );
          if (imageAdded) {
            break;
          }
        }
      }
    }
  }

  responseName = warehouseCode + " " + auditDate + " " + rowNumber;
  doc.setName(responseName);

  return warehouseCode;
}

async function replaceTextWithImagesInCell(
  cell,
  body,
  columnName,
  columnValue
) {
  if (cell.getText().includes("<<" + columnName + ">>")) {
    body.replaceText("<<" + columnName + ">>", " ");
    if (columnValue != "") {
      if (typeof columnValue != "undefined") {
        const links = columnValue.toString().split(",");
        // Getting fileIds from links
        const fileIds = [];
        for (const link of links) {
          if (link.length < 10) {
            continue;
          }
          fileIds.push(extractFileIdFromFileUrlInArray(link));
        }
        // Appending Images
        for (const fileId of fileIds) {
          const blob = UrlFetchApp.fetch(
            Drive.Files.get(fileId).thumbnailLink.replace(/=s.+/, "=s600")
          ).getBlob();
          cell.appendImage(blob);
        }
      }
    }
    return true;
  } else {
    return false;
  }
}
