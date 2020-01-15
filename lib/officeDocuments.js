const officeDocuments = {
  docx: {
    contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
  },
  pptx: {
    contentType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
  },
  xlsx: {
    contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  }
}

module.exports.getDocumentTypes = () => {
  return Object.keys(officeDocuments)
}

module.exports.getContentType = (officeDocumentType) => {
  return officeDocuments[officeDocumentType]
}

module.exports.getDocumentType = (officeDocumentContentType) => {
  return Object.keys(officeDocuments).find((docType) => officeDocuments[docType].contentType === officeDocumentContentType)
}
