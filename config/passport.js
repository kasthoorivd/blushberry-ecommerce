const passport = require("passport");
const GoogleStrategy = require("passport-google-oauth20").Strategy;
const User = require("../models/user/userModel");

passport.use(
  new GoogleStrategy(
    {
      clientID:     process.env.GOOGLE_CLIENT_ID,
      clientSecret: process.env.GOOGLE_CLIENT_SECRET,
      callbackURL:  "/auth/google/callback",
    },
    async (accessToken, refreshToken, profile, done) => {
      try {
        const email        = profile.emails[0].value;
        const profilePhoto = profile.photos[0].value;
        const googleId     = profile.id;
        const fullName     = profile.displayName;

        const existingUser = await User.findOne({ email });

        if (existingUser) {
          /*
           * Case 1 — user signed up with email+password.
           * They have a password field set, meaning they registered normally.
           * Block Google OAuth and redirect to login with a message.
           */
          if (existingUser.password) {
            return done(null, false, {
              message: 'already_registered',
            });
          }

          /*
           * Case 2 — existing Google user. Update their photo/name in case
           * they changed it in Google, then let them in.
           */
          existingUser.profilePhoto = profilePhoto;
          existingUser.fullName     = fullName;
          await existingUser.save();
          return done(null, existingUser);
        }

        /*
         * Case 3 — brand new user. Create their account.
         */
        const newUser = await User.create({
          fullName,
          email,
          googleId,
          profilePhoto,
          isVerified: true,
        });

        return done(null, newUser);

      } catch (error) {
        return done(error, null);
      }
    }
  )
);

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