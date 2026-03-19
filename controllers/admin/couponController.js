const Coupon = require('../../models/user/couponModel')


const loadCoupons = async (req,res) => {
    try {
        const coupons = await Coupon.find().sort({createdAt: -1}).lean()
        res.render('admin/coupons',{coupons})
    } catch (error) {
        console.error(error)
        res.status(500).send('Error loading coupons')
    }
}

const createCoupon = async(req,res) => {
    try {
        const {code,discountAmount,minOrderAmount,maxUses,expiresAt} = req.body

        const existing = await Coupon.findOne({code:code.toUpperCase()})

        if(existing){
            return res.json({success:false,message:'Coupon code already exists'})
        }

        await Coupon.create({
            code,
            discountAmount,
            minOrderAmount:minOrderAmount || 0,
            maxUses: maxUses || null,
            expiresAt : expiresAt || null
        });

        return res.json({success:true,message:'Coupon created'})
    } catch (error) {
        console.error(error)
        res.status(500).json({success:false,message:'Server error'})
    }
}

const toggleCoupon = async (req,res) =>{
    try {
        const coupon = await Coupon.findById(req.params.id)
        coupon.isActive = !coupon.isActive;
        await coupon.save();
        res.json({success:true,isActive:coupon.isActive})
    } catch (error) {
        res.status(500).json({success:false})
    }
}

const deleteCoupon = async (req,res) => {
    try {
        await Coupon.findByIdAndDelete(req.params.id)
        res.json({success:true})
    } catch (error) {
        res.status(500).json({success:false})
    }
}

module.exports = {
    loadCoupons,
    createCoupon,
    toggleCoupon,
    deleteCoupon
}
