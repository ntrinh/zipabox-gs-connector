/**
This file is part of Zipabox GS Connector.

@Author : Nicolas TRINH
*/

/**
 * Utilities functions
 */

/**
 * get SHA1 hash from google script digest algorithm
 */
function getSHA1(digest) {
  var txtHash = '';
  
  for (j = 0; j <digest.length; j++) {
    var hashVal = digest[j];
    if (hashVal < 0) hashVal += 256; 
    if (hashVal.toString(16).length == 1) txtHash += "0";
    
    txtHash += hashVal.toString(16);
  }
  
  return txtHash;
}

/**
 * Check if a JSON object is empty
 */
function isEmpty(obj) {
    return Object.keys(obj).length === 0;
}

