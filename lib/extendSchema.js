const extend = require('node.extend.without.arrays')
const schema = {
  preview: {
    enabled: { type: 'boolean', default: true },
    publicUri: { type: 'string' },
    showWarning: { type: 'boolean', default: true }
  }
}

module.exports = (recipe, recipeSchema) => {
  return {
    office: schema,
    extensions: {
      [recipe]: extend(true, {}, schema, recipeSchema)
    }
  }
}
