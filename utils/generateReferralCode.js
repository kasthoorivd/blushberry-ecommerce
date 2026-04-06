const crypto = require('crypto')

function generateReferralCode(name = '') {
    const prefix = name.slice(0,3).toUpperCase().replace(/[^A-Z]/g ,'X' || 'USR')
      const suffix = crypto.randomBytes(3).toString('hex').toUpperCase()
   return prefix + suffix
}

module.exports = generateReferralCode