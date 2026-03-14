const mongoose = require('mongoose')

const variantSchema = new mongoose.Schema({
    shade:        { type: String, required: true, trim: true },
    varientPrice: { type: Number, required: true, min: 0 },
    salePrice:    { type: Number, default: 0,     min: 0 },
    stock:        { type: Number, required: true, default: 0, min: 0 },
    image:        { type: String, default: '' }    // cloudinary URL — shade swatch
})

const productSchema = new mongoose.Schema(
    {
        name: {
            type:     String,
            required: true,
            trim:     true
        },
        description: {
            type:     String,
            required: true,
            trim:     true
        },
        categoryId: {
            type:     mongoose.Schema.Types.ObjectId,
            ref:      'Category',
            required: true
        },

        // main product gallery — minimum 3 images recommended
        images: {
            type:     [String],    // cloudinary URLs
            validate: {
                validator: function(arr) { return arr.length >= 1 },
                message:   'At least one product image is required'
            }
        },

        // product level offer — category offer is handled separately
        offer: {
            type:    Number,
            default: 0,
            min:     0,
            max:     100
        },

        variants: {
            type:     [variantSchema],
            validate: {
                validator: function(arr) { return arr.length >= 1 },
                message:   'At least one shade variant is required'
            }
        },

        isListed:  { type: Boolean, default: true  },
        isDeleted: { type: Boolean, default: false }
    },
    { timestamps: true }
)

module.exports = mongoose.model('Product', productSchema)