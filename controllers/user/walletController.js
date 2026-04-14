const Wallet = require('../../models/user/walletModel')
const Order = require('../../models/user/orderModel')

async function getOrCreateWallet(userId){
    let wallet = await Wallet.findOne({userId})
    if(!wallet) {
        wallet = await Wallet.create({userId,balance: 0 , transactions : []})
    }
    return wallet
}

const getWallet = async (req,res) =>{
    try {
        const wallet = await getOrCreateWallet(req.session.user._id)

        const transactions = [...wallet.transactions].sort((a,b) => new Date(b.createdAt) - new Date(a.createdAt))


        res.render('user/wallet' ,{
            user:req.session.user,
            wallet,
            transactions,
            success:req.query.success || null,
            error:req.query.error || null
        })
    } catch (error) {
        console.error('getWallet error : ' , error)
        res.redirect('/profile')
    }
}

const creditWallet = async (userId,amount ,description, orderId= null) =>{
    const wallet = await getOrCreateWallet(userId)
    wallet.balance +=amount;
    wallet.transactions.push({
        type:'credit',
        amount,
        description,
        orderId,
        status:'completed'
    })
    await wallet.save()
    return wallet
}


const debitWallet = async (userId, amount, description , orderId= null) =>{
    const wallet = await getOrCreateWallet(userId)

    if(wallet.balance < amount ){
        throw new Error('Insufficient wallet balance')
    }

    wallet.balance -= amount ;
    wallet.transactions.push({
        type:'debit',
        amount,
        description,
        orderId,
        status:'completed'
    })

    await wallet.save()
    return wallet
}

const getWalletBalance = async(userId) =>{
    const wallet = await getOrCreateWallet(userId)
    return wallet.balance
}

const handleCancellationRefund = async (order,userId) =>{
    if(order.paymentMethod === 'COD' && order.paymentStatus !== 'paid') return

    const refundAmount = order.finalAmount 
    const description = `Refund for cancelled order #${order.orderId || order._id}`

    await creditWallet(userId.toString(),refundAmount,description,order._id)
}

const handleReturnRefund = async(order) =>{
    if(!order.userId) throw new Error('order has no userId')

    if(order.paymentMethod === 'COD' && order.paymentStatus !== 'paid') return;

    const refundAmount = order.finalAmount 
    const description = `Refund for returned order #${order.orderId || order._id}`


    await creditWallet(order.userId.toString(), refundAmount , description , order._id)
}

const  applyWalletPayment = async (userId, orderTotal) => {
  const wallet = await getOrCreateWallet(userId);
  const walletUsed = Math.min(wallet.balance, orderTotal);
  const remainingDue = orderTotal - walletUsed;
 
  return { walletUsed, remainingDue, walletBalance: wallet.balance };
};

module.exports = {
    getWallet,
    creditWallet,
    debitWallet,
    getWalletBalance,
    handleCancellationRefund,
    handleReturnRefund,
    applyWalletPayment
}