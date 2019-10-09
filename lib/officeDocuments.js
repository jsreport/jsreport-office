
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
  if (officeDocuments[officeDocumentType] == null) {
    return undefined
  }

  return officeDocuments[officeDocumentType].contentType
}

module.exports.getDocumentType = (officeDocumentContentType) => {
  const types = Object.keys(officeDocuments)

  const found = types.find((docType) => {
    return officeDocuments[docType].contentType === officeDocumentContentType
  })

  if (!found) {
    return undefined
  }

  return found
}
