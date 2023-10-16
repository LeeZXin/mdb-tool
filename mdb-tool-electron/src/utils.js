const fs = require('fs')

exports.readFile = (filePath) => {
    return new Promise((resolve, reject) => {
        fs.readFile(filePath, {
            encoding: 'utf-8'
        }, (err, data) => {
            if (err) {
                reject(err);
            } else {
                resolve(data);
            }
        });
    });
}
