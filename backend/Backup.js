/* ============================================================
   BACKUP
   ============================================================ */

function backupDatabase(token) {
  const username = requireAuth_(token);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const file = DriveApp.getFileById(ss.getId());
  const folder = DriveApp.getFileById(file.getId()).getParents().hasNext() ? DriveApp.getFileById(file.getId()).getParents().next() : DriveApp.getRootFolder();
  const stamp = Utilities.formatDate(new Date(),Session.getScriptTimeZone(),"yyyyMMdd-HHmmss");
  const copy = file.makeCopy("BACKUP_CV_PANDS_TEKSTIL_"+stamp,folder);
  writeActivityLog_(username,"Backup Database","Sistem",copy.getId(),copy.getName());
  return {success:true,fileId:copy.getId(),fileName:copy.getName(),fileUrl:copy.getUrl()};
}


