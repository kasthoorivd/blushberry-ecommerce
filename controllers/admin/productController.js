const Product = require('../../models/user/productModel')
const Category = require('../../models/user/categoryModel')
const { cloudinary } = require('../../config/cloudinary')

const LIMIT = 5


async function uploadIfBase64(value, folder) {
  if (!value) return ''
  if (value.startsWith('data:')) {
    const result = await cloudinary.uploader.upload(value, { folder })
    return result.secure_url
  }
  return value   
}

const loadProducts = async (req, res) => {
  try {
    const page = Math.max(1, parseInt(req.query.page) || 1);
    const searchQuery = (req.query.search || '').trim();
    const filter = { isDeleted: false };

    if (searchQuery) {
      filter.name = { $regex: searchQuery, $options: 'i' };
    }

    const totalProducts = await Product.countDocuments(filter);
    const totalPages = Math.ceil(totalProducts / LIMIT);

    const products = await Product.find(filter)
      .populate('categoryId', 'name')
      .sort({ createdAt: -1 })
      .skip((page - 1) * LIMIT)
      .limit(LIMIT)
      .lean();
   
    
    res.render('admin/products', {
      products,
      currentPage: page,
      totalPages,
      totalProducts,
      limit: LIMIT,
      searchQuery,
      user: req.session.admin || null
    });

    
  } catch (err) {
    console.error('getProducts error:', err);
    res.status(500).render('error', { message: 'Could not load products.' });
  }
}

const loadAddProduct = async (req, res) => {
  try {
    const categories = await Category.find({ isDeleted: false }).lean();
    res.render('admin/addProduct', { categories });
  } catch (err) {
    console.error('getAddProduct error:', err);
    res.status(500).render('error', { message: 'Could not load form.' });
  }
}

const addProduct = async (req, res) => {
  try {
    const { name, description, categoryId, offer, images, variants } = req.body

    const uploadedImages = await Promise.all(
      images.map(base64 => cloudinary.uploader.upload(base64, { folder: 'blushberry/products' }))
    )
    const imageUrls = uploadedImages.map(r => r.secure_url)

   
    const processedVariants = await Promise.all(
      variants.map(async (v) => {
       
        const swatchUrl = await uploadIfBase64(v.image, 'blushberry/shades')

        const shadeGalleryUrls = await Promise.all(
          (v.images || []).map(img => uploadIfBase64(img, 'blushberry/shades'))
        )

        return {
          shade:        v.shade,
          varientPrice: v.varientPrice,
          salePrice:    v.salePrice || 0,
          stock:        v.stock,
          image:        swatchUrl,
          images:       shadeGalleryUrls
        }
      })
    )

    const product = new Product({
      name,
      description,
      categoryId,
      offer: offer || 0,
      images: imageUrls,
      variants: processedVariants
    })

    await product.save()
    return res.json({ success: true, message: 'Product added successfully' })

  } catch (error) {
    console.error('addProduct error:', error)
    return res.json({ success: false, message: 'Failed to add product' })
  }
}

const loadEditProduct = async (req, res) => {
  try {
    const product = await Product.findOne({ _id: req.params.id, isDeleted: false })
    if (!product) return res.status(400).render('error', { message: 'Product not found' })

    const categories = await Category.find({ isDeleted: false }).lean();
    res.render('admin/editProduct', { product, categories })
  } catch (error) {
    console.error('loadEditProduct error:', error)
    res.status(500).render('error', { message: 'Could not load product' })
  }
}

const editProduct = async (req, res) => {
  try {
    const { id } = req.params
    const { name, description, categoryId, offer, existingImages, newImages, variants } = req.body

  
    const uploadedNew = await Promise.all(
      (newImages || []).map(base64 =>
        cloudinary.uploader.upload(base64, { folder: 'blushberry/products' })
      )
    )
    const finalImages = [
      ...(existingImages || []),
      ...uploadedNew.map(r => r.secure_url)
    ]

    
    const processedVariants = await Promise.all(
      variants.map(async (v) => {
       
        const swatchUrl = await uploadIfBase64(v.image, 'blushberry/shades')

      
        const kept = v.keptImages || []
        const uploaded = await Promise.all(
          (v.newImages || []).map(img => uploadIfBase64(img, 'blushberry/shades'))
        )
        const finalShadeGallery = [...kept, ...uploaded]

        return {
          shade:        v.shade,
          varientPrice: v.varientPrice,
          salePrice:    v.salePrice || 0,
          stock:        v.stock,
          image:        swatchUrl,
          images:       finalShadeGallery
        }
      })
    )

    await Product.findByIdAndUpdate(id, {
      name,
      description,
      categoryId,
      offer: offer || 0,
      images: finalImages,
      variants: processedVariants
    })

    return res.json({ success: true, message: 'Product updated successfully' })

  } catch (error) {
    console.error('editProduct error:', error)
    return res.json({ success: false, message: 'Failed to update product' })
  }
}

const toggleProductListing = async (req, res) => {
  try {
    const product = await Product.findOne({ _id: req.params.id, isDeleted: false });
    if (!product) return res.status(404).json({ success: false, message: 'Product not found.' });

    product.isListed = !product.isListed;
    await product.save();

    res.json({
      success: true,
      isListed: product.isListed,
      message: `Product ${product.isListed ? 'listed' : 'unlisted'} successfully.`,
    });
  } catch (err) {
    console.error('toggleProductListing error:', err);
    res.status(500).json({ success: false, message: 'Toggle failed.' });
  }
};

const deleteProduct = async (req, res) => {
  try {
    const product = await Product.findById(req.params.id);
    if (!product) return res.status(404).json({ success: false, message: 'Product not found.' });

    if (product.isDeleted) {
      return res.status(400).json({ success: false, message: 'Product already deleted.' });
    }

    product.isDeleted = true;
    product.isListed = false;
    await product.save();

    res.json({ success: true, message: `"${product.name}" has been deleted.` });
  } catch (err) {
    console.error('deleteProduct error:', err);
    res.status(500).json({ success: false, message: 'Delete failed.' });
  }
};

module.exports = {
  loadProducts,
  loadAddProduct,
  addProduct,
  loadEditProduct,
  editProduct,
  toggleProductListing,
  deleteProduct
}