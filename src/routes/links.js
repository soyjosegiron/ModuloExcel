const express = require('express');
const router = express.Router();
const multer = require('multer');
const path = require('path'); 
const sqlsp = require('sqlstoreprocedure');
const pool = new sqlsp('userAgentsTEST','157.245.89.79','smartDriverAgentsTEST','(vu9bWnA3t&"');
const { dirname } = require('path/posix');
var XLSX = require('xlsx')

//configuracion de storage

const storage = multer.diskStorage(
    {
        destination:(req, file, cb) =>{
            cb(null, path.join(__dirname,'../archivos'))
        }, 

        filename: (req, file, cb) =>{
            cb(null, file.originalname)
        }
    }
);


const tempfile = multer({storage:storage});

router.get('/add', async (req, res) => {
    res.render('links/add');
});

router.post('/add', tempfile.single('excel'),(req, res)=>{
        console.log(req.file);    
        res.send('Recibido');
    if(req.file){
        const excel = XLSX.readFile(path.join(__dirname, '../archivos/'+req.file.originalname));
        var nombreHoja = excel.SheetNames;
        let datos = XLSX.utils.sheet_to_json(excel.Sheets[nombreHoja[0]]);
        //console.log(datos); 
       // console.log(datos.length);
        const datosFormato = datos.map((d)=>{
            return({
                    "IdAgent":d.IdAgent,
                    "Name":d.Name,
                    "Phone":d.Phone,
                    "City":d.City,
                    "Colonia":d.Colonia,
                    "Referencia":d.Referencia,
                    "Compania":d.Compania,
                    "MondayIn":d.MondayIn,
                    "MondayOut":d.MondayOut,
                    "TuesdayIn":d.TuesdayIn,
                    "TuesdayOut":d.TuesdayOut,
                    "WednesdayIn":d.WednesdayIn,
                    "WednesdayOut":d.WednesdayOut,
                    "ThursdayIn":d.ThursdayIn,
                    "ThursdayOut":d.ThursdayOut,
                    "FridayIn":d.FridayIn,
                    "FridayOut":d.FridayOut,
                    "SaturdayIn":d.SaturdayIn,
                    "SaturdayOut":d.SaturdayOut,
                    "SundayIn":d.SundayIn,
                    "SundayOut":d.SundayOut,
                    "Comentario":d.Comentario
            })
        });
        
        datosFormato.forEach(async X =>{
            let repuesta = await(pool.exec('SP_LOAD_AGENT_DATA',{
                agentEmployeeId: X.IdAgent, 
                agentFullname : X.Name,
                agentPhone : X.Phone,
                townName : X.City, 
                neighborhoodName: X.Colonia,
                agentReferencePoint : X.Referencia,
                countName : X.Compania,
                mondayIn : X.MondayIn, 
                mondayOut: X.MondayOut, 
                tuesdayIn: X.TuesdayIn ,
                tuesdayOut:X.TuesdayOut, 
                wednesdayIn: X.WednesdayIn,
                wednesdayOut: X.WednesdayOut, 
                thursdayIn: X.ThursdayIn,
                thursdayOut: X.ThursdayOut, 
                fridayIn: X.FridayIn, 
                fridayOut: X.FridayOut,
                saturdayIn: X.SaturdayIn, 
                saturdayOut: X.SaturdayOut, 
                sundayIn: X.SundayIn, 
                sundayOut: X.SundayOut, 
                agentComment: X.Comentario
              } )  );
              console.log(repuesta.recordsets);
        
        });
        
    }


    
});



module.exports = router;