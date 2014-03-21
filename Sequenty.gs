/**
 * An extremely simple synchronous sequential processing module for node
 * https://github.com/AndyShin/sequenty
 *
 * Modified to be used with Google Script
 * @Author: Nicolas Trinh
 */
function Sequenty() {
  this.run = function(funcs) {
    var i = 0;	
    var recursive = function()
    {      
      funcs[i](function() 
               {
                 i++;
                 
                 if (i < funcs.length)
                   recursive();
               }, i);
    };
    
    recursive();	
  }
}

