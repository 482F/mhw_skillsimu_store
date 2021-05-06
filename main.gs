function onEdit(e){
  var sheet = e.source.getActiveSheet();
  var e_range = e.range;
  var rowStart = e_range.rowStart;
  var rowEnd = e_range.rowEnd;
  var columnStart = e_range.columnStart;
  var columnEnd = e_range.columnEnd;
  if (sheet.getName() == "所持スロ4珠" || sheet.getName() == "非所持スロ4珠" || sheet.getName() == "所持スロ3以下珠"){
    return;
  } else if (sheet.getName() == "装飾品finder"){
    if (rowStart != 1 || rowEnd != 1 || columnStart != 1 || columnEnd != 1){
      return;
    }
    var dataThreeSheet = e.source.getSheetByName("所持スロ3以下珠");
    var dataFourSheet = e.source.getSheetByName("所持スロ4珠");
    accessoriesFinderEdit(sheet,dataThreeSheet, dataFourSheet);
    return;
  }
  if (rowStart == 2 && rowEnd == 2 && columnStart != 1){
    for (var c = columnStart; c <= columnEnd; c += 1){
      generateSkillsAndWeaponSlotFromURL(sheet, sheet.getRange(2, c));
    }
  }else if (rowStart != 1 && rowStart != 2 && columnStart != 1){
    for (var c = columnStart; c <= columnEnd; c += 1){
      generateURLFromSkillsAndWeaponSlot(sheet, sheet.getRange(3, c));
    }
  }
}
function accessoriesFinderEdit(sheet, dataThreeSheet, dataFourSheet){
  sheet.getRange(4, 1, sheet.getMaxRows()-1, 4).clear({contentsOnly: true, skipFilteredRows: true});
  var accessories = sheet.getRange(1, 1).getValue()
  var pattern = /((.+?【.】)\*(\d))/;
  var allPattern = /((.+?【.】)\*(\d))/g;
  var ownAccessories = dataFourSheet.getRange(2, 1, dataFourSheet.getMaxRows()-1, 3).getValues().map(ownAccessory => {
    return [4].concat(ownAccessory);
  }).concat(dataThreeSheet.getRange(2, 1, dataThreeSheet.getMaxRows()-1, 3).getValues().map(ownAccessory => {
    return ownAccessory.concat([100]);
  }));
  accessoryValues = accessories.match(allPattern).map(accessory => {
    var accessoryMatch = accessory.match(pattern);
    var accessoryIndex = ownAccessories.findIndex(acce => acce[2] == accessoryMatch[2]);
    var ownAccessoryNumber = 100;
    if (accessoryIndex != -1){
      ownAccessoryNumber = ownAccessories[accessoryIndex][3];
    }
    var accessoryIndexInOwn = ownAccessories[accessoryIndex][1];
    //[size, name, number, index, ownNumber]
    return [ownAccessories[accessoryIndex][0], accessoryMatch[2], accessoryMatch[3], accessoryIndexInOwn, ownAccessoryNumber];
  });
  for (var i = 0; i < accessoryValues.length; i++){
    var accessoryValue = accessoryValues[i];
    var accessorySize = accessoryValue[0];
    var accessoryRequiredNumber = accessoryValue[2];
    var accessoryIndex = accessoryValue[3];
    var ownAccessoryNumber = accessoryValue[4];
    // 必要数が所持数以上でなければスキップ
    if (accessoryRequiredNumber < ownAccessoryNumber){
      continue;
    }
    accessoryValues = accessoryValues.map((mapAccessoryValue, index) => {
      // i 以降の行に対して、同じサイズで i の珠より後に表示されるものだった場合、インデックスを 1 下げる
      if (i <= index && accessoryIndex < mapAccessoryValue[3] && accessorySize == mapAccessoryValue[0]){
        mapAccessoryValue[3] -= 1;
      }
      return mapAccessoryValue;
    })
  }
  accessoryValues = accessoryValues.map(mapAccessoryValue => {
    var accessoryIndex = mapAccessoryValue[3];
    var pageIndex = "";
    var inPageIndex = "";
    if (accessoryIndex != -1){
      pageIndex = Math.floor(accessoryIndex / 12) + 1;
      inPageIndex = (accessoryIndex % 12) + 1;
      if (6 < inPageIndex){
        inPageIndex -= 13;
      }
    }
    // [name, number, pageIndex, inPageIndex]
    console.log(mapAccessoryValue);
    return [mapAccessoryValue[1], mapAccessoryValue[2], pageIndex, inPageIndex];
  });
  sheet.getRange(4, 1, accessoryValues.length, 4).setValues(accessoryValues);
  return;
}
function generateSkillsAndWeaponSlotFromURL(sheet, targetCell){
  sheet.getRange(3, targetCell.getColumn(), sheet.getMaxRows()-2, 1).clear({contentsOnly: true, skipFilteredRows: true});
  
  var URL = decodeURI(targetCell.getValue());
  var skillsStr = URL.replace(/^https:\/\/mhw\.wiki-db\.com\/sim\/\#skills\=|\&.+$/g, "");
  var weaponSkill = URL.replace(/(^.*\&ws\=|\&d\=.*$)/g, "");
  var weaponSlot = URL.replace(/(^.*\&w\=|\&ws\=.*$)/g, "");
  var skills = skillsStr.split("%2C");
  skills = skills.map(function(value){return [value];});
  var skillsAndWeaponSlot = [[weaponSlot]].concat([[weaponSkill]]);
  skillsAndWeaponSlot = skillsAndWeaponSlot.concat(skills);
  var skillsAndWeaponSlotCells = targetCell.offset(1, 0, skillsAndWeaponSlot.length, 1);
  skillsAndWeaponSlotCells.setValues(skillsAndWeaponSlot);
}

function generateURLFromSkillsAndWeaponSlot(sheet, targetCell){
  var weaponSlot = targetCell.getValue();
  var weaponSkill = targetCell.offset(1, 0).getValue();
  var skillsStartCell = targetCell.offset(2, 0);
  var skillsStartRow = skillsStartCell.getRow();
  var skillsColumn = skillsStartCell.getColumn();
  var skills = sheet.getRange(skillsStartCell.getRow(), skillsStartCell.getColumn(), sheet.getMaxRows()-3, 1).getValues();
  skills = skills.filter(function(value){return Boolean(value[0]);});
  skills = skills.join("%2C");
  var URLCell = targetCell.offset(-1, 0);
  
  if (skills == "" && weaponSlot == ""){
    URLCell.setValue("");
    return;
  }
  
  var URL = "https://mhw.wiki-db.com/sim/#skills="
  URL += skills;
  URL += "&s=1&e=1&v=0&g=13&w=" + weaponSlot;
  URL += "&ws=" + weaponSkill;
  URL += "&d=0&rf=-100&rw=-100&rt=-100&ri=-100&rd=-100&l=200";
  URLCell.setValue(URL);
}
