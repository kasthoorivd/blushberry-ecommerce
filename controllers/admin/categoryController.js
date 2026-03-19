const User     = require('../../models/user/userModel')
const Category = require('../../models/user/categoryModel')
const statusCode = require('../../utils/statusCode')
const Product  = require('../../models/user/productModel')  // add at top

const loadCategory = async (req, res) => {
    try {
        const page        = parseInt(req.query.page)  || 1
        const limit       = parseInt(req.query.limit) || 2
        const searchQuery = req.query.search          || ''

        const filter = { isDeleted: false }
        if (searchQuery) {
            filter.name = { $regex: new RegExp(searchQuery, 'i') }
        }

        const totalCategories = await Category.countDocuments(filter)
        const totalPages      = Math.ceil(totalCategories / limit)

        const categories = await Category.find(filter)
            .sort({ createdAt: -1 })
            .skip((page - 1) * limit)
            .limit(limit)
            .lean()  

  
        const categoryIds = categories.map(c => c._id)

        const productCounts = await Product.aggregate([
            { $match: { categoryId: { $in: categoryIds }, isDeleted: false } },
            { $group: { _id: '$categoryId', count: { $sum: 1 } } }
        ])

        const countMap = {}
        productCounts.forEach(item => {
            countMap[item._id.toString()] = item.count
        })

        
        categories.forEach(cat => {
            cat.productCount = countMap[cat._id.toString()] || 0
        })

        const user = await User.findById(req.session.userId)

        res.render('admin/category', {
            categories,
            currentPage: page,
            totalPages,
            totalCategories,
            limit,
            searchQuery,
            user
        })

    } catch (err) {
        console.error('loadCategory error:', err)
        res.status(500).send('Server Error')
    }
}


const loadAddCategory = async (req, res) => {
    try {
        const user = await User.findById(req.session.userId)
        res.render('admin/addCategory', { user })
    } catch (err) {
        console.error('loadAddCategory error:', err)
        res.status(500).send('Server Error')
    }
}


const loadEditCategory = async (req, res) => {
    try {
        const category = await Category.findById(req.params.id)
        if (!category) return res.redirect('/admin/category')

        const user = await User.findById(req.session.userId)
        res.render('admin/editCategory', { category, user })
    } catch (err) {
        console.error('loadEditCategory error:', err)
        res.status(500).send('Server Error')
    }
}


const addCategory = async (req, res) => {
    try {
        const { name, description, offer } = req.body

        if (!name || !name.trim()) {
            return res.status(400).json({ success: false, message: 'Category name is required' })
        }
        if (!description || !description.trim()) {
            return res.status(400).json({ success: false, message: 'Description is required' })
        }

        const existingCategory = await Category.findOne({
            name: { $regex: new RegExp(`^${name.trim()}$`, 'i') }
        })
        if (existingCategory) {
            return res.status(409).json({ success: false, message: 'Category already exists' })
        }

        const newCategory = new Category({
            name:        name.trim(),
            description: description.trim(),
            offer:       offer ? Number(offer) : 0,
            isListed:    true,
            isDeleted:   false
        })
        await newCategory.save()

        return res.status(201).json({ success: true, message: 'Category added successfully' })

    } catch (error) {
        console.error('addCategory error:', error)
        return res.status(500).json({ success: false, message: 'Something went wrong. Please try again.' })
    }
}


const editCategory = async (req, res) => {
    try {
        const categoryId = req.params.id
        const { name, description, offer } = req.body

        const category = await Category.findById(categoryId)
        if (!category) {
            return res.status(404).json({ success: false, message: 'Category not found' })
        }

        const existingCategory = await Category.findOne({
            name: { $regex: new RegExp(`^${name.trim()}$`, 'i') },
            _id:  { $ne: categoryId }
        })
        if (existingCategory) {
            return res.status(409).json({ success: false, message: 'Category already exists' })
        }

        await Category.findByIdAndUpdate(
            categoryId,
            {
                name:        name.trim(),
                description: description.trim(),
                offer:       Number(offer) || 0
            },
            { new: true }
        )

        return res.status(200).json({ success: true, message: 'Category updated successfully' })

    } catch (error) {
        console.error('editCategory error:', error)
        return res.status(500).json({ success: false, message: 'Something went wrong. Please try again.' })
    }
}


const deleteCategory = async (req, res) => {
    try {
        const categoryId = req.params.id

        const category = await Category.findById(categoryId)
        if (!category) {
            return res.status(404).json({ success: false, message: 'Category not found' })
        }

        if (category.isDeleted) {
            return res.status(400).json({ success: false, message: 'Category already deleted' })
        }

        await Category.findByIdAndUpdate(
            categoryId,
            { isDeleted: true, isListed: false },
            { new: true }
        )

        return res.status(200).json({ success: true, message: 'Category deleted successfully' })

    } catch (error) {
        console.error('deleteCategory error:', error)
        return res.status(500).json({ success: false, message: 'Something went wrong. Please try again.' })
    }
}




const toggleCategoryListing = async (req, res) => {
    try {
        const categoryId = req.params.id

        const category = await Category.findById(categoryId)
        if (!category) {
            return res.status(404).json({
                success: false,
                message: 'Category not found'
            })
        }

        const newListingStatus = !category.isListed
        category.isListed = newListingStatus
        await category.save()

        await Product.updateMany(
            { categoryId: categoryId },
            { isListed: newListingStatus }
        )

        return res.status(200).json({
            success: true,
            message: newListingStatus ? 'Category listed successfully' : 'Category unlisted successfully'
        })

    } catch (error) {
        console.error('toggleCategoryListing error:', error)
        return res.status(500).json({
            success: false,
            message: 'Something went wrong. Please try again.'
        })
    }
}
module.exports = {
    loadCategory,
    loadAddCategory,
    loadEditCategory,
    addCategory,
    editCategory,
    deleteCategory,
    toggleCategoryListing
}