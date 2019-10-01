const axios = require('axios')
const FormData = require('form-data')
const toArray = require('stream-to-array')
const { promisify } = require('util')
const officeDocuments = require('./officeDocuments')
const toArrayAsync = promisify(toArray)

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

  const officeDocumentContentType = officeDocuments.getContentType(officeDocumentType)

  if (!officeDocumentContentType) {
    throw new Error(`Unsupported office document type "${officeDocumentType}"`)
  }

  if (enabled === false || (req.options.preview !== true && req.options.preview !== 'true')) {
    res.meta.fileExtension = officeDocumentType
    res.meta.contentType = officeDocumentContentType
    return
  }

  const publicUriTarget = publicUri || 'http://jsreport.net/temp'

  res.meta.officeDocumentContent = res.content
  res.meta.previewPublicUri = publicUriTarget

  const form = new FormData()
  form.append('field', res.content, `file.${officeDocumentType}`)
  const resp = await axios.post(publicUriTarget, form, {
    headers: form.getHeaders()
  })

  const iframe = '<iframe style="height:100%;width:100%" src="https://view.officeapps.live.com/op/view.aspx?src=' +
    encodeURIComponent((publicUriTarget + '/') + resp.data) + '" />'
  const html = `<html><head><title>${res.meta.reportName}</title><body>` + iframe + '</body></html>'
  res.content = Buffer.from(html)
  res.meta.contentType = 'text/html'
  res.meta.fileExtension = officeDocumentType
}
