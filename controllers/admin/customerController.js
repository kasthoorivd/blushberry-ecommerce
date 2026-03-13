const User = require('../../models/user/userModel')

const loadCustomer = async (req, res) => {
  try {
    const page = parseInt(req.query.page) || 1;
    const limit = parseInt(req.query.limit) || 5;
    const search = req.query.search?.trim() || "";
    const sortField = req.query.sortField || "createdAt";
    const sortOrder = req.query.sortOrder || "desc";
    const skip = (page - 1) * limit;

    const searchFilter = search
      ? {
        $or: [
          { fullName: { $regex: search, $options: "i" } },
          { email: { $regex: search, $options: "i" } },
          { phoneNumber: { $regex: search, $options: "i" } },
        ],
      }
      : {};

    const totalUsers = await User.countDocuments(searchFilter);
    const totalPages = Math.ceil(totalUsers / limit);


    const sortOptions = { [sortField]: sortOrder === "asc" ? 1 : -1 };

    const customers = await User.find(searchFilter)
      .sort(sortOptions)
      .skip(skip)
      .limit(limit)
      .select("-password");

    return res.render("admin/customers", {
       user: req.session.user || null,
      customers,
      currentPage: page,
      totalPages,
      totalUsers,
      limit,
      search,
      sortField,
      sortOrder,
      hasNextPage: page < totalPages,
      hasPrevPage: page > 1,
    });
  } catch (error) {
    console.error("loadCustomer error:", error);
    return res.redirect("/admin/dashboard");
  }
};


const blockUser = async (req, res) => {
  try {
    const { id } = req.params;

    const customer = await User.findById(id);
    if (!customer) {
      return res.status(404).json({ success: false, message: "User not found" });
    }

    if (customer.isBlocked) {
      return res.status(400).json({ success: false, message: "User is already blocked" });
    }

    customer.isBlocked = true;
    await customer.save();

    return res.status(200).json({
      success: true,
      message: `${customer.fullName} has been blocked successfully`,
    });
  } catch (error) {
    console.error("blockUser error:", error);
    return res.status(500).json({ success: false, message: "Server error" });
  }
};



const unblockUser = async (req, res) => {
  try {
    const { id } = req.params;

    const customer = await User.findById(id);
    if (!customer) {
      return res.status(404).json({ success: false, message: "User not found" });
    }

    if (!customer.isBlocked) {
      return res.status(400).json({ success: false, message: "User is not blocked" });
    }

    customer.isBlocked = false;
    await customer.save();

    return res.status(200).json({
      success: true,
      message: `${customer.fullName} has been unblocked successfully`,
    });
  } catch (error) {
    console.error("unblockUser error:", error);
    return res.status(500).json({ success: false, message: "Server error" });
  }
};

const loadEditCustosmer = async (req,res) => {
  try {
    const customer = await User.findById(req.params.id).lean()
    if(!customer) return res.redirect('/admin/customers')
    return res.render('admin/editCustomer',{customer,errors:null,success:null})
  } catch (error) {
    console.error('loadEditCustomer error:',error)
    return res.redirect('/admin/customers')
  }
}


const editCustomer = async (req, res) => {
  try {
    const { fullName, email, phoneNumber } = req.body
    const userId = req.params.id

    if (!fullName || !/^[A-Za-z\s]{3,}$/.test(fullName)) {
      return res.json({ success: false, message: 'Enter a valid name (letters only, min 3 chars)' })
    }

    if (!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
      return res.json({ success: false, message: 'Enter a valid email address' })
    }

    if (phoneNumber && !/^\d{10}$/.test(phoneNumber)) {
      return res.json({ success: false, message: 'Enter a valid 10-digit phone number' })
    }

    const existing = await User.findOne({ email, _id: { $ne: userId } })
    if (existing) {
      return res.json({ success: false, message: 'Email already in use by another account' })
    }

    await User.findByIdAndUpdate(userId, {
      fullName,
      email,
      phoneNumber: phoneNumber || null,
    })

    return res.json({ success: true })

  } catch (error) {
    console.error('editCustomer error:', error)
    return res.json({ success: false, message: 'Something went wrong' })
  }
}

const deleteCustomer = async(req,res) =>{
  try {
    await User.findByIdAndDelete(req.params.id)
    return res.json({success:true,message:'Customer deleted successfully'})
  } catch (error) {
    console.log('deleteCustomer error :',error)
    return res.json({success:true,message:'Failed to delete customer'})
  }
}
module.exports = {
  loadCustomer,
  blockUser,
  unblockUser,
  loadEditCustosmer,
  editCustomer,
  deleteCustomer
};