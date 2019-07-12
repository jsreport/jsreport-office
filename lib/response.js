const axios = require('axios')
const FormData = require('form-data')
const toArray = require('stream-to-array')
const { promisify } = require('util')
const toArrayAsync = promisify(toArray)

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

module.exports = async ({
  previewOptions: { publicUri, enabled },
  officeDocumentType,
  stream,
  buffer
}, req, res) => {
  if (buffer) {
    res.content = buffer
  } else {
    res.content = Buffer.concat(await toArrayAsync(stream))
  }

  if (enabled === false || (req.options.preview !== true && req.options.preview !== 'true')) {
    res.meta.fileExtension = officeDocumentType
    res.meta.contentType = officeDocuments[officeDocumentType]
    return
  }

  const form = new FormData()
  form.append('field', res.content, `file.${officeDocumentType}`)
  const resp = await axios.post(publicUri || 'http://jsreport.net/temp', form, {
    headers: form.getHeaders()
  })

  const iframe = '<iframe style="height:100%;width:100%" src="https://view.officeapps.live.com/op/view.aspx?src=' +
    encodeURIComponent((publicUri || 'http://jsreport.net/temp' + '/') + resp.data) + '" />'
  const html = `<html><head><title>${res.meta.reportName}</title><body>` + iframe + '</body></html>'
  res.content = Buffer.from(html)
  res.meta.contentType = 'text/html'
  res.meta.fileExtension = officeDocumentType
}
