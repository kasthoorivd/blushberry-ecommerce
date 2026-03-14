const cloudinary = require('cloudinary').v2;
const { CloudinaryStorage } = require('multer-storage-cloudinary');
const multer = require('multer');

cloudinary.config({
  cloud_name: process.env.CLOUDINARY_CLOUD_NAME,
  api_key:    process.env.CLOUDINARY_API_KEY,
  api_secret: process.env.CLOUDINARY_API_SECRET,
});

// Storage for main product gallery images
const productStorage = new CloudinaryStorage({
  cloudinary,
  params: {
    folder:         'blushberry/products',
    allowed_formats: ['jpg', 'jpeg', 'png', 'webp'],
    transformation: [{ width: 800, height: 800, crop: 'limit', quality: 'auto' }],
  },
});

// Storage for shade swatch images (variants)
const swatchStorage = new CloudinaryStorage({
  cloudinary,
  params: {
    folder:          'blushberry/swatches',
    allowed_formats: ['jpg', 'jpeg', 'png', 'webp'],
    transformation:  [{ width: 200, height: 200, crop: 'fill', quality: 'auto' }],
  },
});

const uploadProductImages = multer({ storage: productStorage });
const uploadSwatchImages  = multer({ storage: swatchStorage  });

module.exports = { cloudinary, uploadProductImages, uploadSwatchImages };