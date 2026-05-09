const mongoose = require("mongoose")

const connectDb = async () => {

    try {

        await mongoose.connect("mongodb://127.0.0.1:27017/blushberry")

        console.log("MongoDB Connected Successfully")

    }
    catch (error) {

        console.log("MongoDB Connection Failed:", error.message)

        process.exit(1)

    }

}

module.exports = connectDb