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

app.use(nocache())

 app.use(session({
  secret: "yourSecretKey",
  resave: false,
  saveUninitialized: false,
  store: MongoStore.create({
    mongoUrl: process.env.MONGO_URI
  }),
  cookie: {
    maxAge: 1000 * 60 * 60
  }
}));

app.use(passport.initialize())
app.use(passport.session())
app.set('view engine','ejs')

app.set('views',path.join(__dirname,'views')) 

app.use('/',userRouter)
app.use('/admin',adminRouter)

// app.get('/',(req,res)=>{
//     res.send('Blushberry Home')
// })

const PORT = process.env.PORT || 3000;

app.listen(PORT, () => {

    console.log(`Server running on http://localhost:${PORT}`);

});
