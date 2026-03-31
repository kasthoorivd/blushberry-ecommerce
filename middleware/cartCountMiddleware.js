const Cart = require('../models/user/cartModel');

const attachCartCount = async (req, res, next) => {
  try {
    if (req.session?.user?._id) {
      const cart = await Cart.findOne({ userId: req.session.user._id });
    
      res.locals.cartCount = cart
        ? cart.items.reduce((s, i) => s + i.quantity, 0)
        : 0;
    } else {
      res.locals.cartCount = 0;
    }
  
  } catch (err) {
   
    res.locals.cartCount = 0;
  }
  next();
};

module.exports = attachCartCount;