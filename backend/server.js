var express = require("express");
var app = express();
var db = require("./database.js");
var bodyParser = require("body-parser");
const { request, response } = require("express");
const res = require("express/lib/response");
app.use(bodyParser.json());
const md5 = require('md5')

let HTTP_PORT = 3000;

const cors = require('cors');
app.use(cors({
    origin: "*"
}));


app.listen(HTTP_PORT, () => {
    console.log("Server is running on %PORT%".replace("%PORT%", HTTP_PORT))
});

//home page
app.get('/', function (req, res, next) {
    res.render('server', { title: 'Express', session: req.session });
});

//singin
app.post("/singing", function (req, res, next) {

    try {
        var errors = []

        if (!req.body) {
            errors.push("An invalid input");
        }
        const { fullname, email, mobile, password } = req.body;

        const hash_password = md5(password.toString());

        //check email adress
        const checkEmail = `SELECT email FROM user WHERE email=?`;
        db.query(checkEmail, [email], (err, result, fields) => {
            if (!result.length) {
                const sql = `INSERT INTO user(fullname, email, mobile, password) VALUES (?,?,?,?)`;
                db.query(sql, [fullname, email, mobile, hash_password], (err, result, fields) => {
                    if (err) {
                        res.status(400).json({
                            "error": err.message,
                            "message": "Error1 ",
                        })
                        return;
                    } else {
                        console.log(result.userid);
                        res.status(201).json({
                            "message": "Singup success! " + req.body.fullname ,
                            "status": 200,
                            
                        })
                    }
                })
            } else {
                res.json({
                    "message": "Email address alrady used ",
                })
            }
        });

    } catch (E) {
        res.status(400).send(E);
    }
});

//login
app.post("/login", function (req, res, next) {

    try {
        var errors = []

        if (!req.body) {
            errors.push("An invalid input");
        }
        const { email, password } = req.body;

        if (email && password) {

            const checkEmail = `SELECT * FROM user WHERE email = ?`;
            db.query(checkEmail, [email], (error, data, result) => {

                if (error) throw error;

                
                if (data.length > 0) {
                    if (data[0].password == md5(password)) {
                        res.status(201).json({
                            "message": "loging success! " ,
                            "status": 200,
                            "userId" :data[0].userid,
                            "userName" : data[0].fullname
      
                        });
                    } else {
                        res.json({
                            "message": "Wrong Password ",
                            "status": 201
                        });
                    }

                } else {
                    res.json({
                        "message": "Wrong Email Address",
                    });
                }
            });
            return;
        }
        else {
            res.json({
                "message": "Please Enter Email Address and Password Details",
            });
            return;
        }

    } catch (E) {

        res.status(400).send(E);
    }
});

//add new project

app.post("/addnew", function (req, res, next) {

    try {
        var errors = []

        if (!req.body) {
            errors.push("An invalid input");
        }
        const { Name, Location, Description, SubDate, Reminder, ProjectType, SurveyHelpers, model} = req.body;
        
        //check project name
        const checkProjectName = `SELECT Name FROM project WHERE Name=?`;
        db.query(checkProjectName, [Name], (err, result, fields) => {
            if (!result.length) {
                const sql = `INSERT INTO project(Name, Location, Description, SubDate, Reminder,ProjectType, SurveyHelpers, model) VALUES (?,?,?,?,?,?,?,?)`;

                db.query(sql, [ Name, Location, Description, SubDate, Reminder, ProjectType, SurveyHelpers, model], (err, result, fields) => {
                    if (err) {
                        res.status(400).json({
                            "error": err.message,
                            "message": "Error1 ",
                            "status": 201
                        })
                        return;
                    } else {
                        res.json({
                            "message": "New Project Added!" + req.body.Name,
                            "status": 200
                        })
                    }
                });
            } else {
                res.json({
                    "message": "Project name alredy used!",
                    "status": 201
                })
            }
        });
    } catch (E) {
        res.status(400).send(E);
    }
});

//project list
app.get("/projects", function (req, res, next) {

    try {
        var errors = []

        if (!req.body) {
            errors.push("An invalid input");
        }
        
        const sql = `SELECT Name, date_format(CreatedDate, '%Y-%m-%d %H:%i')date FROM project`;

        db.query(sql, (err, result, data) => {
            if (err) {
                res.status(400).json({
                    "error": err.message,
                    "message": "No projects",
                    "status": 201
                })
                return;
            } else {
                res.status(200).json(result);
            }
        });


    } catch (E) {
        res.status(400).send(E);
    }
});

//update project

app.post("/editProject", function (req, res, next) {

    try {
        var errors = []

        if (!req.body) {
            errors.push("An invalid input");
        }
        const { Name, Location, Description, SubDate, Reminder, ProjectType, SurveyHelpers, model, ProjectId } = req.body;
        const sql = `UPDATE project SET Name =?, Location=?, Description=?, SubDate=?, Reminder=?,ProjectType=?, SurveyHelpers =?, model=?) WHERE ProjectId=?`;

        db.query(sql, [Name, Location, Description, SubDate, Reminder, ProjectType, SurveyHelpers, model, ProjectId], (err, result, fields) => {
            if (err) {
                res.status(400).json({
                    "error": err.message,
                    "message": "Somthing went wrong",
                    "status": 201
                })
                return;
            } else {
                res.json({
                    "message": "Successfully update " + req.body.Name,
                    "status": 200
                })
            }
        });


    } catch (E) {
        res.status(400).send(E);
    }
});

//delete project
app.post("/deleteProject", function (req, res, next) {

    try {
        var errors = []

        if (!req.body) {
            errors.push("An invalid input");
        }
        const { ProjectId } = req.body;
        const sql = `DELETE FROM project  WHERE ProjectId = ?`;

        db.query(sql, [ProjectId], (err, result, fields) => {
            if (err) {
                res.status(400).json({
                    "error": err.message,
                    "message": "Somthing went wrong",
                    "status": 201
                })
                return;
            } else {
                res.json({
                    "message": "Successfully Deleted Project " + req.body.Name,
                    "status": 200
                })
            }
        });


    } catch (E) {
        res.status(400).send(E);
    }
});

//Create a Report

app.get("/CreateReport" , function(req, res){
    try{
        var errors = []

        if (!req.body) {
            errors.push("An invalid input");
        }
        const { Name} = req.body;
        const sql = `SELECT Name, Description, Location, ProjectType, date_format(SubDate, '%Y-%m-%d %H:%i')sdate,date_format(CreatedDate, '%Y-%m-%d %H:%i')cdate, model, SurveyHelpers FROM project WHERE SubDate <current_date()`;

        db.query(sql, [Name], (err, result, fields) => {
            if (err) {
                res.status(400).json({
                    "error": err.message,
                    "message": "Still not complete the project",
                    "status": 201
                })
                return;
            } else {
                res.status(200).json(result)
            }
        });


    }
    catch(E) {
        res.status(400).send(E);
    }
})

module.exports = app;