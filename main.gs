function onEdit(e){
  var sheet = e.source.getActiveSheet();
  var e_range = e.range;
  var rowStart = e_range.rowStart;
  var rowEnd = e_range.rowEnd;
  var columnStart = e_range.columnStart;
  var columnEnd = e_range.columnEnd;
  if (sheet.getName() == "所持スロ4珠" || sheet.getName() == "非所持スロ4珠"){
    return;
  } else if (sheet.getName() == "装飾品finder"){
    if (rowStart != 1 || rowEnd != 1 || columnStart != 1 || columnEnd != 1){
      return;
    }
    var dataSheet = e.source.getSheetByName("所持スロ4珠");
    accessoriesFinderEdit(sheet,dataSheet);
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
function accessoriesFinderEdit(sheet, dataSheet){
  sheet.getRange(2, 1, sheet.getMaxRows()-1, 4).clear({contentsOnly: true, skipFilteredRows: true});
  var accessories = sheet.getRange(1, 1).getValue()
  var pattern = /((.+?【.】)\*(\d))/;
  var allPattern = /((.+?【.】)\*(\d))/g;
  var ownAccessories = dataSheet.getRange(2, 2, dataSheet.getMaxRows()-1, 2).getValues();
  accessoryValues = accessories.match(allPattern).map(accessory => {
    var accessoryMatch = accessory.match(pattern);
    var accessoryIndex = ownAccessories.findIndex(acce => acce[0] == accessoryMatch[2]);
    var ownAccessoryNumber = 100;
    if (accessoryIndex != -1){
      ownAccessoryNumber = ownAccessories[accessoryIndex][1];
    }
    return [accessoryMatch[2], accessoryMatch[3], accessoryIndex, ownAccessoryNumber];
  });
  for (var i = 0; i < accessoryValues.length; i++){
    var accessoryValue = accessoryValues[i];
    var accessoryRequiredNumber = accessoryValue[1];
    var accessoryIndex = accessoryValue[2];
    var ownAccessoryNumber = accessoryValue[3];
    if (accessoryRequiredNumber < ownAccessoryNumber){
      continue;
    }
    accessoryValues = accessoryValues.map((mapAccessoryValue, index) => {
      if (i < index && accessoryIndex <= mapAccessoryValue[2]){
        mapAccessoryValue[2] -= 1;
      }
      return mapAccessoryValue;
    })
  }
  accessoryValues = accessoryValues.map(mapAccessoryValue => {
    var accessoryIndex = mapAccessoryValue[2];
    var pageIndex = "";
    var inPageIndex = "";
    if (accessoryIndex != -1){
      pageIndex = Math.floor(accessoryIndex / 12) + 1;
      inPageIndex = (accessoryIndex % 12) + 1;
      if (6 < inPageIndex){
        inPageIndex -= 13;
      }
    }
    mapAccessoryValue[2] = pageIndex;
    mapAccessoryValue[3] = inPageIndex;
    return mapAccessoryValue;
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
