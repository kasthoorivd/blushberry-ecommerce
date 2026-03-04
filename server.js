const  express = require('express')
const app = express() 
require('dotenv').config();
const morgan = require('morgan')
const session = require('express-session')
const passport = require('./config/passport')
const MongoStore = require('connect-mongo');


const path = require('path') 
const nocache = require('nocache')

const connectDb = require('./config/connectDb')
app.use(morgan('dev'))

const userRouter = require('./routes/userRoutes')
const adminRouter = require('./routes/adminRoutes')

connectDb()


app.use(express.json())
app.use(express.urlencoded({extended:true}))

app.use(express.static(path.join(__dirname,'public')))


app.use(session({
  secret: "yourSecretKey",
  resave: false,
  saveUninitialized: false,
  store: MongoStore.create({
    mongoUrl: process.env.MONGO_URI
  }),
  cookie: {
    maxAge: 1000 * 60 * 60,
    secure:false,
    httpOnly:true,
    sameSite:'lax'
  }
}));

app.use(passport.initialize())
app.use(passport.session())
app.use(nocache())

// app.get("/debug", (req, res) => {
//   console.log("Session object:", req.session);
//   console.log("req.user:", req.user);
//   console.log("isAuthenticated:", req.isAuthenticated());
//   res.json({
//     session: req.session,
//     user: req.user,
//     isAuthenticated: req.isAuthenticated()
//   });
// });

// app.js or server.js
app.use((req, res, next) => {
  res.locals.user = req.session.user || null;
  next();
});

app.set('view engine','ejs')

app.set('views',path.join(__dirname,'views')) 

app.use('/',userRouter)
app.use('/admin',adminRouter)

// app.get('/',(req,res)=>{
//     res.send('Blushberry Home')
// })
app.get("/check-session", (req, res) => {
  res.json({
    isAuthenticated: req.isAuthenticated(),
    user: req.user,
    session: req.session
  });
});
const PORT = process.env.PORT || 3000;

app.listen(PORT, () => {

    console.log(`Server running on http://localhost:${PORT}`);

});
