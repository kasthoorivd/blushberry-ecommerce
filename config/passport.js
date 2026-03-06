const passport = require("passport");
const GoogleStrategy = require("passport-google-oauth20").Strategy;
const User = require("../models/user/userModel");

passport.use(
  new GoogleStrategy(
    {
      clientID: process.env.GOOGLE_CLIENT_ID,
      clientSecret: process.env.GOOGLE_CLIENT_SECRET,
      callbackURL: "/auth/google/callback",
    },
 async (accessToken, refreshToken, profile, done) => {
  try {
    const email = profile.emails[0].value;
    const profilePhoto = profile.photos[0].value;
    const googleId = profile.id;
    const fullName = profile.displayName;

    let user = await User.findOneAndUpdate(
      { email },
      { googleId, profilePhoto, fullName },
      { returnDocument:'after'}
    );

    if (!user) {
      user = await User.create({
        fullName,
        email,
        googleId,
        profilePhoto,
        isVerified: true,
      });
    }

    return done(null, user);
  } catch (error) {
    return done(error, null);
  }
}
  ))
passport.serializeUser((user, done) => {
  done(null, user._id);
});

passport.deserializeUser(async (id, done) => {
  try {
    const user = await User.findById(id);
    done(null, user);
  } catch (error) {
    done(error, null);
  }
});

module.exports = passport;