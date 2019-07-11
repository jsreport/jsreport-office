const archiver = require('archiver')
const fs = require('fs')

module.exports = async ({ reporter, files, officeDocumenType }, req, res) => {
  const {
    pathToFile: officeFileName,
    stream: output
  } = await reporter.writeTempFileStream((uuid) => `${uuid}.${officeDocumenType}`)

  await new Promise((resolve, reject) => {
    const archive = archiver('zip')

    output.on('close', () => {
      reporter.logger.debug('Successfully zipped now.', req)
      res.stream = fs.createReadStream(officeFileName)
      resolve()
    })

    archive.on('error', (err) => reject(err))

    archive.pipe(output)

    files.forEach((f) => archive.append(f.data, { name: f.path }))

    archive.finalize()
  })
}
