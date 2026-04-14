const User = require('../../models/user/userModel')
const Otp = require('../../models/user/otpModel')
const Address = require('../../models/user/addressModel')
const Coupon = require('../../models/user/couponModel')
const Wallet = require('../../models/user/walletModel')
const bcrypt = require('bcrypt')
const crypto = require('crypto')
const mongoose = require('mongoose')
const generateOtp = require('../../utils/generateOtp')
const sendEmail = require('../../utils/sendEmail')

// ── Generate referral code for users who don't have one ──
const generateReferralCode = (fullName) => {
    const prefix = fullName.substring(0, 3).toUpperCase();
    const random = crypto.randomBytes(3).toString('hex').toUpperCase();
    return `${prefix}${random}`;
};


const loadProfile = async (req, res) => {
    try {
        const userId = req.session.user._id;
        const user = await User.findById(userId).lean();
        const address = await Address.findOne({ user: userId, isDefault: true });
        const success = req.query.success === 'true' ? 'Profile updated successfully!' : null;
        const pwSuccess = req.query.pwSuccess === 'true';
        return res.render('user/userProfile', { user, address: address || null, success, errors: null, pwSuccess })
    } catch (error) {
        res.status(500).json({ error: 'page not loading' })
    }
}


const updateProfile = async (req, res) => {
    try {
        const findUser = await User.findById(req.session.user._id).lean();

        if (!findUser) {
            return res.render('user/userProfile', { error: "User not found", user: findUser });
        }

        const { fullName, phoneNumber } = req.body;
        const errors = {};

        const nameRegex = /^[A-Za-z\s]{3,}$/;
        if (!fullName) {
            errors.fullName = "Full name is required";
        } else if (!nameRegex.test(fullName)) {
            errors.fullName = "Enter a valid full name (only letters, min 3 chars)";
        }

        const phoneRegex = /^\d{10}$/;
        if (!phoneNumber) {
            errors.phoneNumber = "Phone number is required";
        } else if (!phoneRegex.test(phoneNumber)) {
            errors.phoneNumber = "Enter a valid 10-digit phone number";
        }

        if (Object.keys(errors).length > 0) {
            return res.render('user/userProfile', {
                errors,
                user: findUser,
                success: null,
                address: null,
                pwSuccess: false
            });
        }

        const updateData = { fullName, phoneNumber };
        if (req.file) {
            updateData.profilePhoto = '/uploads/profiles/' + req.file.filename;
        }

        const updatedUser = await User.findByIdAndUpdate(
            req.session.user._id,
            updateData,
            { new: true }
        ).lean();

        // Keep profilePhoto in session so header avatar doesn't break
        req.session.user = {
            _id: updatedUser._id,
            fullName: updatedUser.fullName,
            email: updatedUser.email,
            isBlocked: updatedUser.isBlocked,
            profilePhoto: updatedUser.profilePhoto
        };

        res.redirect('/profile');

    } catch (error) {
        console.log(error);
        res.status(500).json({ error: "cannot update" });
    }
}


const changePassword = async (req, res) => {
    try {
        const { currentPassword, newPassword, confirmPassword } = req.body;
        const errors = {};

        const userId = new mongoose.Types.ObjectId(req.session.user._id);
        const findUser = await User.findById(userId);
        if (!findUser) return res.redirect('/profile');

        const passwordMatch = await bcrypt.compare(currentPassword, findUser.password);

        if (!currentPassword) {
            errors.currentPassword = 'Current password is required';
        } else if (!passwordMatch) {
            errors.currentPassword = 'Invalid current password';
        }

        const passwordRegex = /^(?=.*[a-z])(?=.*[A-Z])(?=.*\d)(?=.*[@#$%^&*!])[A-Za-z\d@#$%^&*!]{8,}$/;
        if (!newPassword) {
            errors.newPassword = 'Password is required';
        } else if (!passwordRegex.test(newPassword)) {
            errors.newPassword = 'Min 8 chars with uppercase, lowercase, number & special character';
        }

        if (!confirmPassword) {
            errors.confirmPassword = 'Please confirm your password';
        } else if (newPassword !== confirmPassword) {
            errors.confirmPassword = 'Passwords do not match';
        }

        if (Object.keys(errors).length > 0) {
            const address = await Address.findOne({ user: findUser._id, isDefault: true });
            return res.render('user/userProfile', {
                errors, user: findUser, success: null, address: address || null, pwSuccess: false
            });
        }

        const passwordHash = await bcrypt.hash(newPassword, 10);
        const result = await User.findByIdAndUpdate(
            findUser._id,
            { $set: { password: passwordHash } },
            { new: true }
        );
        console.log('Password updated for:', result?.email);

        res.redirect('/profile?pwSuccess=true');

    } catch (error) {
        console.log('ERROR:', error);
        res.redirect('/profile');
    }
};


const requestEmailChange = async (req, res) => {
    try {
        const { newEmail } = req.body;
        const userId = req.session.user?._id;

        if (!userId) return res.json({ success: false, message: "Not logged in" });
        if (!newEmail) return res.json({ success: false, message: "New email is required" });

        const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
        if (!emailRegex.test(newEmail)) {
            return res.json({ success: false, message: "Enter a valid email address" });
        }

        const existing = await User.findOne({ email: newEmail });
        if (existing) return res.json({ success: false, message: "Email already in use" });

        const otp = generateOtp();
        const expiresAt = Date.now() + 5 * 60 * 1000;

        await Otp.deleteMany({ email: newEmail, purpose: 'emailChange' });
        await Otp.create({ email: newEmail, otp, expiresAt, purpose: 'emailChange' });

        req.session.pendingEmail = newEmail;
        await sendEmail(newEmail, otp);

        return res.json({ success: true, redirectUrl: '/otp' });

    } catch (err) {
        console.error('requestEmailChange error:', err);
        res.status(500).json({ success: false, message: "Server error" });
    }
};


const loadCoupons = async (req, res) => {
    try {
        const userId = req.session.user?._id;
        const now = new Date();

        const availableCoupons = await Coupon.find({
            isActive: true,
            $or: [{ expiresAt: null }, { expiresAt: { $gt: now } }]
        }).lean();

        const coupons = availableCoupons.map(c => {
            const usedByUser = c.usedBy.some(id => id.toString() === userId.toString());
            const isExhausted = c.maxUses && c.usedBy.length >= c.maxUses;
            let status = 'active';
            if (usedByUser) status = 'used';
            else if (isExhausted) status = 'expired';
            return { ...c, status };
        });

        return res.render('user/coupon', { coupons });

    } catch (error) {
        console.log(error);
        res.status(500).json({ error: 'Could not load coupons' });
    }
};


// ── Referral info API ──
// Auto-generates a referral code for old users who signed up before this feature.
// Falls back to req.protocol + host if BASE_URL env var is not set.
const getReferralInfo = async (req, res) => {
    try {
        const userId = req.session.user?._id;
        if (!userId) return res.status(401).json({ success: false, message: 'Not logged in' });

        let user = await User.findById(userId).lean();

        // ── Generate code on the fly for users who signed up before referrals existed ──
        if (!user.referralCode) {
            const newCode = generateReferralCode(user.fullName);
            user = await User.findByIdAndUpdate(
                userId,
                { $set: { referralCode: newCode } },
                { new: true }
            ).lean();
            console.log(`Generated missing referral code for ${user.email}: ${user.referralCode}`);
        }

        const wallet        = await Wallet.findOne({ userId }).lean();
        const referralCount = await User.countDocuments({ referredBy: userId });

        // Use BASE_URL from .env — fallback to current host for local dev
        const baseUrl = process.env.BASE_URL
            ? process.env.BASE_URL.replace(/\/$/, '')          // strip trailing slash
            : `${req.protocol}://${req.get('host')}`;

        return res.json({
            success:      true,
            referralCode: user.referralCode,
            referralLink: `${baseUrl}/signup?ref=${user.referralCode}`,
            referralCount,
            walletBalance: wallet?.balance || 0
        });
    } catch (err) {
        console.error('getReferralInfo error:', err);
        return res.status(500).json({ success: false, message: 'Server error' });
    }
};


module.exports = {
    loadProfile,
    updateProfile,
    changePassword,
    requestEmailChange,
    loadCoupons,
    getReferralInfo
}