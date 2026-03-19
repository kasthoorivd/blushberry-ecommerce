const Product = require('../../models/user/productModel')
const Review = require('../../models/user/reviewModel')


function attachDisplayFields(p){
    const prices = p.variants.map(v=>v.salePrice > 0 ? v.salePrice : v.varientPrice)
    p.displayPrice = Math.min(...prices)
    p.displayMaxPrice = Math.max(...prices)
    const origPrices = p.variants.map(v=>v.varientPrice)
    p.originalPrice = Math.min(...origPrices)
    p.displayOffer = p.offer || 0
    p.inStock = p.variants.some(v => v.stock > 0)
    return p
}


const loadProductDetail = async (req,res) => {
    try {
        const product = await Product.findOne({
            _id: req.params.id,
            isDeleted : false
        })
        .populate('categoryId','name')
        .lean()


        if(!product || !product.isListed){
            req.session.toast = {type:'error',message:'This product is no longer available'}
            return res.redirect('/products')
        }

        attachDisplayFields(product)

        const reviews = await Review.find({productId:product._id, isDeleted:false})
        .populate('userId','fullName profilePhoto')
        .sort({createdAt : -1})
        .lean()


        const avgRating = reviews.length
        ?(reviews.reduce((sum,r)=>sum + r.rating,0)/reviews.length).toFixed(1)
        :0

        const ratingBreakdown = {5:0,4:0,3:0,2:0,1:0}
        reviews.forEach(r=>{ratingBreakdown[r.rating]=(ratingBreakdown[r.rating] || 0)+1 })


        const userId = req.session.user?._id
        const userReview = userId
        ?reviews.find(r => String(r.userId?._id)===String(userId))
        :null

        const productShades = product.variants.map(v => v.shade)
 
    const related = await Product.find({
      _id:       { $ne: product._id },
      isDeleted: false,
      isListed:  true,
      $or: [
        { categoryId: product.categoryId?._id },
        { 'variants.shade': { $in: productShades } }
      ]
    })
    .populate('categoryId', 'name')
    .limit(4)
    .lean()
 
    related.forEach(attachDisplayFields)
 
    res.render('user/productDetail', {
      product,
      reviews,
      avgRating,
      ratingBreakdown,
      reviewCount: reviews.length,
      userReview,
      related,
      user: req.session.user || null
    })

    } catch (error) {
         console.error('loadProductDetail error:', err)
    // invalid ObjectId or DB error → redirect to listing
    return res.redirect('/products')
  }
    
}

const submitReview = async (req,res) => {
    try {
        const userId = req.session.user?._id
        if(!userId){
            return res.status(401).json({success:false,message:'Please log in to leave a review'})
        }

        const { rating ,title , comment } = req.body
        const productId = req.params.id

        const product = await Product.findOne({_id:productId,isDeleted:false,isListed:true})
        if(!product){
            return res.status(404).json({success:false,message:'Product not found'})

        }

  const ratingNum = parseInt(rating)
    if (!ratingNum || ratingNum < 1 || ratingNum > 5) {
      return res.status(400).json({ success: false, message: 'Rating must be between 1 and 5.' })
    }
 
    // upsert — update if exists, create if not
    await Review.findOneAndUpdate(
      { productId, userId },
      { rating: ratingNum, title: (title || '').trim(), comment: (comment || '').trim(), isDeleted: false },
      { upsert: true, new: true }
    )
 
    res.json({ success: true, message: 'Review submitted successfully!' })
 
  } catch (err) {
    console.error('submitReview error:', err)
    res.status(500).json({ success: false, message: 'Could not submit review.' })
  }
}

const deleteReview = async (req, res) => {
  try {
    const userId = req.session.user?._id
    if (!userId) {
      return res.status(401).json({ success: false, message: 'Not logged in.' })
    }
 
    await Review.findOneAndUpdate(
      { productId: req.params.id, userId },
      { isDeleted: true }
    )
 
    res.json({ success: true, message: 'Review deleted.' })
  } catch (err) {
    console.error('deleteReview error:', err)
    res.status(500).json({ success: false, message: 'Could not delete review.' })
  }
}
 


module.exports = {
    loadProductDetail,
    submitReview,
    deleteReview
}