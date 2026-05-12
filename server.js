const express       = require('express')
const app           = express()
require('dotenv').config()
const morgan        = require('morgan')
const path          = require('path')
const session       = require('express-session')
const passport      = require('./config/passport')
const MongoStore    = require('connect-mongo')
const methodOverride = require('method-override')
const nocache       = require('nocache')

const connectDB        = require('./config/connectDB')
const userRouter        = require('./routes/userRoutes')
const adminRouter       = require('./routes/adminRoutes')
const attachCartCount   = require('./middleware/cartCountMiddleware')

connectDB()

app.use(morgan('dev'))
app.use(methodOverride('_method'))
app.use(express.json({ limit: '50mb' }))
app.use(express.urlencoded({ limit: '50mb', extended: true }))
app.use(express.static(path.join(__dirname, 'public')))
app.use(nocache())

app.set('view engine', 'ejs')
app.set('views', path.join(__dirname, 'views'))


const mongoStoreOptions = { mongoUrl: process.env.MONGO_URI }

const cookieBase = {
  maxAge: 1000 * 60 * 60,
  secure: false,
  httpOnly: true,
  sameSite: 'lax'
}


const userSession = session({
  name: 'user.sid',                        
  secret: process.env.USER_SESSION_SECRET || 'user-secret-key',
  resave: false,
  saveUninitialized: false,
  store: MongoStore.create(mongoStoreOptions),
  cookie: cookieBase
})


const adminSession = session({
  name: 'admin.sid',                        
  secret: process.env.ADMIN_SESSION_SECRET || 'admin-secret-key',
  resave: false,
  saveUninitialized: false,
  store: MongoStore.create(mongoStoreOptions),
  cookie: cookieBase
})


app.use('/admin', adminSession)

app.use('/', userSession)
app.use(passport.initialize())
app.use(passport.session())        


app.use(attachCartCount)

app.use((req, res, next) => {
  res.locals.user        = req.session.user || null
  res.locals.currentPath = req.path
  res.locals.success     = null
  res.locals.error       = null
  res.locals.errors      = null
  res.locals.formData    = {}
  next()
})

app.use((req, res, next) => {
  res.status(404).render('error', { 
    statusCode: 404,
    message: 'Page not found' 
  })
})

app.use((err, req, res, next) => {
  console.error(err.stack)
  const statusCode = err.status || err.statusCode || 500
  const message    = err.message || 'Something went wrong'
  res.status(statusCode).render('error', { statusCode, message })
})

app.use('/', userRouter)
app.use('/admin', adminRouter)

const PORT = process.env.PORT || 3000
app.listen(PORT, () => console.log(`Server running on http://localhost:${PORT}`))