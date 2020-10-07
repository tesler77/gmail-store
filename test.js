const fs = require('fs');

fs.readdir('./src/public',(err,res)=>{
    let a = res.find((single)=>{
        return single == 'ccc';        
    })
    console.log(a);
})