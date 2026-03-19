const  express = require('express')
const app = express() 
require('dotenv').config();
const morgan = require('morgan')
const path = require('path')
const session = require('express-session')
const passport = require('./config/passport')
const MongoStore = require('connect-mongo');

const methodOverride = require('method-override')
app.use(methodOverride('_method'))   



const nocache = require('nocache')

const connectDb = require('./config/connectDb')
app.use(morgan('dev'))

const userRouter = require('./routes/userRoutes')
const adminRouter = require('./routes/adminRoutes')

connectDb()

app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ limit: '50mb', extended: true }));
// app.use(express.json())
// app.use(express.urlencoded({extended:true}))

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

app.use((req, res, next) => {
  res.locals.user = req.session.admin || null;
  next();
});

app.use((req, res, next) => {
  res.locals.currentPath = req.path;
  next();
});


app.set('view engine','ejs')

app.set('views',path.join(__dirname,'views')) 

app.use('/',userRouter)
app.use('/admin',adminRouter)


const PORT = process.env.PORT || 3000;

app.listen(PORT, () => {

    console.log(`Server running on http://localhost:${PORT}`);

});