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

const connectDb         = require('./config/connectDb')
const userRouter        = require('./routes/userRoutes')
const adminRouter       = require('./routes/adminRoutes')
const attachCartCount   = require('./middleware/cartCountMiddleware')

connectDb()

app.use(morgan('dev'))
app.use(methodOverride('_method'))
app.use(express.json({ limit: '50mb' }))
app.use(express.urlencoded({ limit: '50mb', extended: true }))
app.use(express.static(path.join(__dirname, 'public')))
app.use(nocache())

app.set('view engine', 'ejs')
app.set('views', path.join(__dirname, 'views'))

// ─── Shared session store (both sessions use same DB, different cookie names) ───
const mongoStoreOptions = { mongoUrl: process.env.MONGO_URI }

const cookieBase = {
  maxAge: 1000 * 60 * 60,
  secure: false,
  httpOnly: true,
  sameSite: 'lax'
}

// ─── User session middleware ───────────────────────────────────────────────────
const userSession = session({
  name: 'user.sid',                          // <-- unique cookie name
  secret: process.env.USER_SESSION_SECRET || 'user-secret-key',
  resave: false,
  saveUninitialized: false,
  store: MongoStore.create(mongoStoreOptions),
  cookie: cookieBase
})

// ─── Admin session middleware ──────────────────────────────────────────────────
const adminSession = session({
  name: 'admin.sid',                         // <-- unique cookie name
  secret: process.env.ADMIN_SESSION_SECRET || 'admin-secret-key',
  resave: false,
  saveUninitialized: false,
  store: MongoStore.create(mongoStoreOptions),
  cookie: cookieBase
})

// ─── Passport only runs in user context ───────────────────────────────────────
app.use('/admin', adminSession)

app.use('/', userSession)
app.use(passport.initialize())
app.use(passport.session())        // passport only touches the user session

// ─── Shared locals ────────────────────────────────────────────────────────────
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

app.use('/', userRouter)
app.use('/admin', adminRouter)

const PORT = process.env.PORT || 3000
app.listen(PORT, () => console.log(`Server running on http://localhost:${PORT}`))