// ==================== Google Drive API Wrapper ====================

const Drive = (() => {

  async function _request(fn) {
    await Auth.ensureToken();
    try {
      return await fn();
    } catch (e) {
      const msg = e?.result?.error?.message || e?.message || 'Drive API error';
      console.error('Drive error:', msg, e);
      throw new Error(msg);
    }
  }

  async function createFolder(name, parentId) {
    return _request(async () => {
      const meta = {
        name,
        mimeType: 'application/vnd.google-apps.folder',
        parents: [parentId || CONFIG.DRIVE_FOLDER_ID]
      };
      const r = await gapi.client.drive.files.create({
        resource: meta,
        fields: 'id, name, webViewLink'
      });
      return r.result;
    });
  }

  async function uploadFile(file, folderId, fileName) {
    await Auth.ensureToken();
    const token = Auth.getToken();
    if (!token) throw new Error('ไม่ได้เข้าสู่ระบบ');

    const parentId = folderId || CONFIG.DRIVE_FOLDER_ID;
    const metadata = {
      name: fileName || file.name,
      parents: [parentId]
    };

    const form = new FormData();
    form.append('metadata', new Blob([JSON.stringify(metadata)], { type: 'application/json' }));
    form.append('file', file);

    const response = await fetch(
      'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,name,webViewLink,webContentLink',
      {
        method: 'POST',
        headers: { Authorization: `Bearer ${token}` },
        body: form
      }
    );

    if (!response.ok) {
      const err = await response.json();
      throw new Error(err.error?.message || 'Upload failed');
    }
    return await response.json();
  }

  async function getFileInfo(fileId) {
    return _request(async () => {
      const r = await gapi.client.drive.files.get({
        fileId,
        fields: 'id, name, webViewLink, webContentLink, mimeType'
      });
      return r.result;
    });
  }

  async function listFilesInFolder(folderId) {
    return _request(async () => {
      const r = await gapi.client.drive.files.list({
        q: `'${folderId}' in parents and trashed=false`,
        fields: 'files(id,name,webViewLink,mimeType,modifiedTime)',
        orderBy: 'createdTime desc'
      });
      return r.result.files || [];
    });
  }

  async function deleteFile(fileId) {
    return _request(async () => {
      await gapi.client.drive.files.delete({ fileId });
    });
  }

  async function createLoanFolder(loanId, memberName) {
    const folderName = `${loanId}_${memberName}`;
    const folder = await createFolder(folderName, CONFIG.DRIVE_FOLDER_ID);
    return folder;
  }

  async function uploadLoanDocument(file, folderId, docType) {
    const fileNames = {
      contract: 'สัญญากู้เงิน',
      id_card: 'สำเนาบัตรประชาชน',
      house_reg: 'สำเนาทะเบียนบ้าน',
      photo: 'รูปถ่าย'
    };
    const ext = file.name.split('.').pop();
    const displayName = fileNames[docType] || docType;
    const fileName = `${displayName}.${ext}`;
    return await uploadFile(file, folderId, fileName);
  }

  async function makeFilePublic(fileId) {
    return _request(async () => {
      await gapi.client.drive.permissions.create({
        fileId,
        resource: { role: 'reader', type: 'anyone' }
      });
    });
  }

  return {
    createFolder,
    uploadFile,
    getFileInfo,
    listFilesInFolder,
    deleteFile,
    createLoanFolder,
    uploadLoanDocument,
    makeFilePublic
  };
})();
