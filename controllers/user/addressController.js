const Address = require('../../models/user/addressModel')


const loadAddresses = async (req, res) => {
  try {
    const user = req.session.user
    const userId = user._id
    const page = parseInt(req.query.page) || 1
    const limit = 2
    const skip = (page - 1) * limit
    const totalAddresses = await Address.countDocuments({user:userId})
    const totalPages = Math.ceil(totalAddresses/limit)
    const addresses = await Address.find({ user: userId })
    .sort({ createdAt: -1 })
    .skip(skip)
    .limit(limit)
    return res.render('user/addresses', { user, addresses, successMsg: null, errorMsg: null ,currentPage : page,totalAddresses,totalPages})
  } catch (error) {
    console.error('loadAddresses error:', error)
    return res.redirect('/profile')
  }
}


const loadAddAddress = (req, res) => {
  const user = req.session.user
  return res.render('user/addAddress', { user, errorMsg: null })
}


const addAddress = async (req, res) => {
  try {
    const userId = req.session.user._id

    const {
      name, country, state, address, city,
      pincode, addressType, landmark,
      mobile, email, alternateNumber,
    } = req.body

   
    const addresses = await Address.find({ user: userId }).sort({ createdAt: -1 })

    if (!name || !country || !state || !address || !pincode || !mobile || !email) {
      return res.render('user/addresses', {
        user: req.session.user,
        addresses,
        errorMsg: 'Please fill in all required fields',
        successMsg: null,
      })
    }

    if (!/^\d{6}$/.test(pincode)) {
      return res.render('user/addresses', {
        user: req.session.user,
        addresses,
        errorMsg: 'Pin code must be a 6-digit number',
        successMsg: null,
      })
    }

    if (!/^\d{10}$/.test(mobile)) {
      return res.render('user/addresses', {
        user: req.session.user,
        addresses,
        errorMsg: 'Mobile number must be 10 digits',
        successMsg: null,
      })
    }

    const newAddress = new Address({
      user: userId,
      name,
      country,
      state,
      address,
      city: city || null,
      pincode,
      addressType: addressType || 'Home',
      landmark: landmark || null,
      mobile,
      email,
      alternateNumber: alternateNumber || null,
    })

    await newAddress.save()

    return res.redirect('/addresses?success=true')

  } catch (error) {
    console.error('addAddress error:', error)
    return res.redirect('/addresses')
  }
}



const loadEditAddress = async (req, res) => {
  try {
    const userId = req.session.user._id
    const addressId = req.params.id
    const page = parseInt(req.query.page) || 1
    const limit = 2
    const skip = (page - 1) * limit
    const totalAddresses = await Address.countDocuments({user:userId})
    const totalPages = Math.ceil(totalAddresses/limit)
    const address = await Address.findOne({ _id: addressId, user: userId })
    if (!address) return res.redirect('/addresses')
    return res.render('user/addresses', {
   user: req.session.user, 
   errorMsg: null,
   currentPage:page,
   totalAddresses,
   totalPages
   })
  } catch (error) {
    console.error('loadEditAddress error:', error)
    return res.redirect('/addresses')
  }
}



const editAddress = async (req, res) => {
  try {
    const userId    = req.session.user._id
    const addressId = req.params.id

    const {
      name, country, state, address, city,
      pincode, addressType, landmark,
      mobile, email, alternateNumber,
    } = req.body

    if (!name || !country || !state || !address || !pincode || !mobile || !email) {
      return res.json({ success: false, message: 'Please fill in all required fields' })
    }

    if (!/^\d{6}$/.test(pincode)) {
      return res.json({ success: false, message: 'Pin code must be a 6-digit number' })
    }

    if (!/^\d{10}$/.test(mobile)) {
      return res.json({ success: false, message: 'Mobile number must be 10 digits' })
    }

    await Address.findOneAndUpdate(
      { _id: addressId, user: userId },
      {
        name, country, state, address,
        city: city || null,
        pincode,
        addressType: addressType || 'Home',
        landmark: landmark || null,
        mobile, email,
        alternateNumber: alternateNumber || null,
      },
      { new: true }
    )

    return res.json({ success: true })

  } catch (error) {
    console.error('editAddress error:', error)
    return res.json({ success: false, message: 'Something went wrong' })
  }
}


const deleteAddress = async (req, res) => {
  try {
    const userId    = req.session.user._id
    const addressId = req.params.id

    await Address.findOneAndDelete({ _id: addressId, user: userId })

    return res.json({ success: true })

  } catch (error) {
    console.error('deleteAddress error:', error)
    return res.json({ success: false, message: 'Failed to delete address' })
  }
}


const setDefaultAddress = async (req, res) => {
  try {
    const userId    = req.session.user._id
    const addressId = req.params.id

    await Address.updateMany({ user: userId }, { isDefault: false })
    await Address.findOneAndUpdate(
      { _id: addressId, user: userId },
      { isDefault: true }
    )

    return res.json({ success: true })

  } catch (error) {
    console.error('setDefaultAddress error:', error)
    return res.json({ success: false, message: 'Failed to set default' })
  }
}


module.exports = {
  loadAddresses,
  loadAddAddress,
  addAddress,
  loadEditAddress,
  editAddress,
  deleteAddress,
  setDefaultAddress,
}