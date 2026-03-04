const passport = require("passport");
const LocalStrategy = require("passport-local").Strategy;
const GoogleStrategy = require("passport-google-oauth20").Strategy;
const User = require("../models/user/userModel");
const bcrypt = require("bcrypt");

// LOCAL STRATEGY
passport.use(new LocalStrategy(
  { usernameField: "email" },
  async (email, password, done) => {
    const user = await User.findOne({ email });
    if (!user) return done(null, false);

    // const isMatch = await bcrypt.compare(password, user.password);
   
    // if (!isMatch) return done(null, false);

    return done(null, user);
  }
));

// GOOGLE STRATEGY
passport.use(new GoogleStrategy({
  clientID: process.env.GOOGLE_CLIENT_ID,
  clientSecret: process.env.GOOGLE_CLIENT_SECRET,
  callbackURL: "/auth/google/callback"
},
async (accessToken, refreshToken, profile, done) => {

  try {
    let user = await User.findOne({
      email: profile.emails[0].value,      
    });

    if (!user) {
      user = await User.create({
        fullName: profile.displayName,
        email: profile.emails[0].value,
        password: null,
        isVerified: true,
        googleId: profile.id,               // ← store Google ID
        profilePhoto: profile.photos[0].value // ← store Google profile photo
      });
    }

    return done(null, user);

  } catch (err) {
    return done(err, null);
  }
}));

passport.serializeUser((user, done) => {
  done(null, user._id);
});

passport.deserializeUser(async (id, done) => {
try {
    const user = await User.findById(id);
  done(null, user || false);
} catch (error) {
  done(error,null)
}
});

module.exports = passport;