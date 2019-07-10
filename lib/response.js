const axios = require('axios')
const FormData = require('form-data')

const officeDocuments = {
  docx: {
    contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
  },
  pptx: {
    contentType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
  }
}

module.exports = async (options, officeDocumentType, req, res) => {
  if (!req.options.preview || options.preview.enabled === false) {
    res.meta.fileExtension = officeDocumentType
    res.meta.contentType = officeDocuments[officeDocumentType]
    return
  }

  const form = new FormData()
  form.append('field', res.content, 'file.docx')
  const resp = await axios.post(options.preview.publicUri || 'http://jsreport.net/temp', form, {
    headers: form.getHeaders()
  })

  const iframe = '<iframe style="height:100%;width:100%" src="https://view.officeapps.live.com/op/view.aspx?src=' +
    encodeURIComponent((options.preview.publicUri || 'http://jsreport.net/temp' + '/') + resp.data) + '" />'
  const html = '<html><head><title>jsreport</title><body>' + iframe + '</body></html>'
  res.content = Buffer.from(html)
  res.meta.contentType = 'text/html'
  res.meta.fileExtension = officeDocumentType
}
