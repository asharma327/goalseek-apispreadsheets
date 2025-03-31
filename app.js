/************************************************************
 *  app.js — Node/Express example translating the Python code
 *  Run:  node app.js
 *  Then POST to: http://localhost:3000/api/calculate
 ************************************************************/
const express = require("express");
const bodyParser = require("body-parser");

// Now define an Express server with one route:
const app = express();
app.use(bodyParser.json());

app.get("/", (req, res) => {
    res.status(200).json({'msg': 'works'});
})

// Example POST route
app.post("/api/calculate", (req, res) => {
    try {
        // Expect JSON body with:
        //  {
        //    "inputCells": { "15yrlump": { "Xinput_datumvandaag": "2025-03-18", ... } },
        //    "outputCells": { "15yrlump": ["F31","H50","F84"] },
        //    "preFormulasActions": [{ "type": "macro", "parameters": {"name":"lump15yrls"}}],
        //    "postFormulasActions": []
        //  }
        // console.log("Request received")
        let { inputCells, outputCells, preFormulasActions, postFormulasActions } = req.body;
        // console.log(inputCells)
        // console.log(outputCells)
        // console.log(preFormulasActions)
        // console.log(postFormulasActions)
        const { PdZppHJFxW4lidr7 } = require("./workbook");
        let workbook = new PdZppHJFxW4lidr7(inputCells || {});
        let results = workbook.calculateOutputCells(
            outputCells || {},
            preFormulasActions || [],
            postFormulasActions || []
        );

        // console.log(results)

        // let actualResults = {"15yrlump": {"F31": 4800.0, "H50": 73828.48000000004, "H51": 4800.0, "F84": 141028.48000000004, "F86": "Yes"}}
        // console.log(actualResults)

        // Return results as JSON
        res.status(200).json(results);
    } catch (err) {
        console.error("Error in /api/calculate:", err);
        res.status(500).json({ status: "error", message: err.toString() });
    }
});

// Start listening
// const PORT = 3000;
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Server running on port ${PORT}`);
});

