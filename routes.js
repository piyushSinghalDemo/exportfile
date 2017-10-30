// routes.js
const portfolio = require('./models/portfolio');
module.exports = [];

module.exports = [
    {
        method: 'GET',
        path: '/',
        handler: function(request, reply) {
            console.log("In Routes Files...");
            portfolio.find(function(error, res) {
                console.log("response :"+res);
                if (error) {
                    console.error(error);
                }
                reply(res);
            });
        }
    }
  ];