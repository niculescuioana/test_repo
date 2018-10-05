var lib = { 
    'settings': {
      'currency': {
        'symbol' : "$",		// default currency symbol is '$'
        'format' : "%s%v",	// controls output: %s = symbol, %v = value (can be object, see docs)
        'decimal' : ".",		// decimal point separator
        'thousand' : ",",		// thousands separator
        'precision' : 2,		// decimal places
        'grouping' : 3		// digit grouping (not implemented yet)
      },
      'number': {
        'precision' : 0,		// default precision on numbers is 0
        'grouping' : 3,		// digit grouping (not implemented yet)
        'thousand' : ",",
        'decimal' : "."
      }
    }
  };
  
  function main() {
    Logger.log(formatToCurrency('DKK', 1000000));
  }
  
  function formatToCurrency(code,value) {
    var currency = getCurrencyMap(code);
    
    var defaultCurrency = {
      'symbol': '',
      'thousandsSeparator': ',',
      'decimalSeparator': '.',
      'symbolOnLeft': true,
      'spaceBetweenAmountAndSymbol': false,
      'decimalDigits': 2
    }
    
    var formatMapping = [
      {
        symbolOnLeft: true,
        spaceBetweenAmountAndSymbol: false,
        format: {
          pos: '%s%v',
          neg: '-%s%v',
          zero: '%s%v'
        }
      },
      {
        symbolOnLeft: true,
        spaceBetweenAmountAndSymbol: true,
        format: {
          pos: '%s %v',
          neg: '-%s %v',
          zero: '%s %v'
        }
      },
      {
        symbolOnLeft: false,
        spaceBetweenAmountAndSymbol: false,
        format: {
          pos: '%v%s',
          neg: '-%v%s',
          zero: '%v%s'
        }
      },
      {
        symbolOnLeft: false,
        spaceBetweenAmountAndSymbol: true,
        format: {
          pos: '%v %s',
          neg: '-%v %s',
          zero: '%v %s'
        }
      }
    ];
    
    var symbolOnLeft = currency.symbolOnLeft;
    var spaceBetweenAmountAndSymbol = currency.spaceBetweenAmountAndSymbol;
    
    var format = formatMapping.filter(function(f) {
      return f.symbolOnLeft == symbolOnLeft && f.spaceBetweenAmountAndSymbol == spaceBetweenAmountAndSymbol
    })[0].format;
    
    return formatMoney(value, {
      'symbol': currency.symbol,
      'decimal': currency.decimalSeparator,
      'thousand': currency.thousandsSeparator,
      'precision': currency.decimalDigits,
      'format': format
    });
  }
  
  function formatMoney(number, symbol) {
    var opts = defaults(symbol,lib.settings.currency),
        
        // Check format (returns object with pos, neg and zero):
        formats = checkCurrencyFormat(opts.format),
        
        // Choose which format to use for this value:
        useFormat = number > 0 ? formats.pos : number < 0 ? formats.neg : formats.zero;
    
    // Return with currency symbol added:
    return useFormat.replace('%s', opts.symbol).replace('%v', formatNumber(Math.abs(number), checkPrecision(opts.precision), opts.thousand, opts.decimal));
  };
  
  
  function formatNumber(number, precision, thousand, decimal) {
    // Build options object from second param (if object) or all params, extending defaults:
    var opts = defaults(
      (isObject(precision) ? precision : {
        precision : precision,
        thousand : thousand,
        decimal : decimal
      }),
      lib.settings.number
    ),
        
        // Clean up precision
        usePrecision = checkPrecision(opts.precision),
        
        // Do some calc:
        negative = number < 0 ? "-" : "",
          base = parseInt(toFixed(Math.abs(number || 0), usePrecision), 10) + "",
            mod = base.length > 3 ? base.length % 3 : 0;
    
    // Format the number:
    return negative + (mod ? base.substr(0, mod) + opts.thousand : "") + base.substr(mod).replace(/(\d{3})(?=\d)/g, "$1" + opts.thousand) + (usePrecision ? opts.decimal + toFixed(Math.abs(number), usePrecision).split('.')[1] : "");
  };
  
  function checkPrecision(val, base) {
    val = Math.round(Math.abs(val));
    return isNaN(val)? base : val;
  }
  
  function toFixed(value, precision) {
    precision = checkPrecision(precision, lib.settings.number.precision);
    
    var exponentialForm = Number(value + 'e' + precision);
    var rounded = Math.round(exponentialForm);
    var finalResult = Number(rounded + 'e-' + precision).toFixed(precision);
    return finalResult;
  };
  
  function defaults(object, defs) {
    var key;
    object = object || {};
    defs = defs || {};
    // Iterate over object non-prototype properties:
    for (key in defs) {
      if (defs.hasOwnProperty(key)) {
        // Replace values with defaults only if undefined (allow empty/zero values):
        if (object[key] == null) object[key] = defs[key];
      }
    }
    return object;
  }
  
  function checkCurrencyFormat(format) {
    var defaults = "%s%v";
    
    // Allow function as format parameter (should return string or object):
    if ( typeof format === "function" ) format = format();
    
    // Format can be a string, in which case `value` ("%v") must be present:
    if ( isString( format ) && format.match("%v") ) {
      
      // Create and return positive, negative and zero formats:
      return {
        pos : format,
        neg : format.replace("-", "").replace("%v", "-%v"),
        zero : format
      };
      
      // If no format, or object is missing valid positive value, use defaults:
    } else if ( !format || !format.pos || !format.pos.match("%v") ) {
      
      // If defaults is a string, casts it to an object for faster checking next time:
      return ( !isString( defaults ) ) ? defaults : lib.settings.currency.format = {
        pos : defaults,
        neg : defaults.replace("%v", "-%v"),
        zero : defaults
      };
      
    }
    // Otherwise, assume format was fine:
    return format;
  }
  
  function getCurrencyMap(code) {
    return {
      "AED": {
        "code": "AED",
        "symbol": "د.إ.‏",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": true,
        "decimalDigits": 2
      },
      "AFN": {
        "code": "AFN",
        "symbol": "؋",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "ALL": {
        "code": "ALL",
        "symbol": "Lek",
        "thousandsSeparator": ".",
        "decimalSeparator": ",",
        "symbolOnLeft": false,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "AMD": {
        "code": "AMD",
        "symbol": "֏",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": false,
        "spaceBetweenAmountAndSymbol": true,
        "decimalDigits": 2
      },
      "ANG": {
        "code": "ANG",
        "symbol": "ƒ",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "AOA": {
        "code": "AOA",
        "symbol": "Kz",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "ARS": {
        "code": "ARS",
        "symbol": "$",
        "thousandsSeparator": ".",
        "decimalSeparator": ",",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": true,
        "decimalDigits": 2
      },
      "AUD": {
        "code": "AUD",
        "symbol": "$",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "AWG": {
        "code": "AWG",
        "symbol": "ƒ",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "AZN": {
        "code": "AZN",
        "symbol": "₼",
        "thousandsSeparator": " ",
        "decimalSeparator": ",",
        "symbolOnLeft": false,
        "spaceBetweenAmountAndSymbol": true,
        "decimalDigits": 2
      },
      "BAM": {
        "code": "BAM",
        "symbol": "КМ",
        "thousandsSeparator": ".",
        "decimalSeparator": ",",
        "symbolOnLeft": false,
        "spaceBetweenAmountAndSymbol": true,
        "decimalDigits": 2
      },
      "BBD": {
        "code": "BBD",
        "symbol": "$",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "BDT": {
        "code": "BDT",
        "symbol": "৳",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": true,
        "decimalDigits": 0
      },
      "BGN": {
        "code": "BGN",
        "symbol": "лв.",
        "thousandsSeparator": " ",
        "decimalSeparator": ",",
        "symbolOnLeft": false,
        "spaceBetweenAmountAndSymbol": true,
        "decimalDigits": 2
      },
      "BHD": {
        "code": "BHD",
        "symbol": "د.ب.‏",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": true,
        "decimalDigits": 3
      },
      "BIF": {
        "code": "BIF",
        "symbol": "FBu",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": false,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 0
      },
      "BMD": {
        "code": "BMD",
        "symbol": "$",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "BND": {
        "code": "BND",
        "symbol": "$",
        "thousandsSeparator": ".",
        "decimalSeparator": ",",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 0
      },
      "BOB": {
        "code": "BOB",
        "symbol": "Bs",
        "thousandsSeparator": ".",
        "decimalSeparator": ",",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": true,
        "decimalDigits": 2
      },
      "BRL": {
        "code": "BRL",
        "symbol": "R$",
        "thousandsSeparator": ".",
        "decimalSeparator": ",",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": true,
        "decimalDigits": 2
      },
      "BSD": {
        "code": "BSD",
        "symbol": "$",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "BTC": {
        "code": "BTC",
        "symbol": "Ƀ",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": false,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "BTN": {
        "code": "BTN",
        "symbol": "Nu.",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": true,
        "decimalDigits": 1
      },
      "BWP": {
        "code": "BWP",
        "symbol": "P",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "BYR": {
        "code": "BYR",
        "symbol": "р.",
        "thousandsSeparator": " ",
        "decimalSeparator": ",",
        "symbolOnLeft": false,
        "spaceBetweenAmountAndSymbol": true,
        "decimalDigits": 2
      },
      "BZD": {
        "code": "BZD",
        "symbol": "BZ$",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "CAD": {
        "code": "CAD",
        "symbol": "$",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "CDF": {
        "code": "CDF",
        "symbol": "FC",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": false,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "CHF": {
        "code": "CHF",
        "symbol": "CHF",
        "thousandsSeparator": "'",
        "decimalSeparator": ".",
        "symbolOnLeft": false,
        "spaceBetweenAmountAndSymbol": true,
        "decimalDigits": 2
      },
      "CLP": {
        "code": "CLP",
        "symbol": "$",
        "thousandsSeparator": ".",
        "decimalSeparator": ",",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": true,
        "decimalDigits": 2
      },
      "CNY": {
        "code": "CNY",
        "symbol": "¥",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "COP": {
        "code": "COP",
        "symbol": "$",
        "thousandsSeparator": ".",
        "decimalSeparator": ",",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": true,
        "decimalDigits": 2
      },
      "CRC": {
        "code": "CRC",
        "symbol": "₡",
        "thousandsSeparator": ".",
        "decimalSeparator": ",",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "CUC": {
        "code": "CUC",
        "symbol": "CUC",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "CUP": {
        "code": "CUP",
        "symbol": "$MN",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "CVE": {
        "code": "CVE",
        "symbol": "$",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "CZK": {
        "code": "CZK",
        "symbol": "Kč",
        "thousandsSeparator": " ",
        "decimalSeparator": ",",
        "symbolOnLeft": false,
        "spaceBetweenAmountAndSymbol": true,
        "decimalDigits": 2
      },
      "DJF": {
        "code": "DJF",
        "symbol": "Fdj",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": false,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 0
      },
      "DKK": {
        "code": "DKK",
        "symbol": "kr.",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": false,
        "spaceBetweenAmountAndSymbol": true,
        "decimalDigits": 2
      },
      "DOP": {
        "code": "DOP",
        "symbol": "RD$",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "DZD": {
        "code": "DZD",
        "symbol": "د.ج.‏",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": true,
        "decimalDigits": 2
      },
      "EGP": {
        "code": "EGP",
        "symbol": "ج.م.‏",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": true,
        "decimalDigits": 2
      },
      "ERN": {
        "code": "ERN",
        "symbol": "Nfk",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": false,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "ETB": {
        "code": "ETB",
        "symbol": "ETB",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "EUR": {
        "code": "EUR",
        "symbol": "€",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": false,
        "spaceBetweenAmountAndSymbol": true,
        "decimalDigits": 2
      },
      "FJD": {
        "code": "FJD",
        "symbol": "$",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "FKP": {
        "code": "FKP",
        "symbol": "£",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "GBP": {
        "code": "GBP",
        "symbol": "£",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "GEL": {
        "code": "GEL",
        "symbol": "Lari",
        "thousandsSeparator": " ",
        "decimalSeparator": ",",
        "symbolOnLeft": false,
        "spaceBetweenAmountAndSymbol": true,
        "decimalDigits": 2
      },
      "GHS": {
        "code": "GHS",
        "symbol": "₵",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "GIP": {
        "code": "GIP",
        "symbol": "£",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "GMD": {
        "code": "GMD",
        "symbol": "D",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": false,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "GNF": {
        "code": "GNF",
        "symbol": "FG",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": false,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 0
      },
      "GTQ": {
        "code": "GTQ",
        "symbol": "Q",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "GYD": {
        "code": "GYD",
        "symbol": "$",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "HKD": {
        "code": "HKD",
        "symbol": "HK$",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "HNL": {
        "code": "HNL",
        "symbol": "L.",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": true,
        "decimalDigits": 2
      },
      "HRK": {
        "code": "HRK",
        "symbol": "kn",
        "thousandsSeparator": ".",
        "decimalSeparator": ",",
        "symbolOnLeft": false,
        "spaceBetweenAmountAndSymbol": true,
        "decimalDigits": 2
      },
      "HTG": {
        "code": "HTG",
        "symbol": "G",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "HUF": {
        "code": "HUF",
        "symbol": "Ft",
        "thousandsSeparator": " ",
        "decimalSeparator": ",",
        "symbolOnLeft": false,
        "spaceBetweenAmountAndSymbol": true,
        "decimalDigits": 2
      },
      "IDR": {
        "code": "IDR",
        "symbol": "Rp",
        "thousandsSeparator": ".",
        "decimalSeparator": ",",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 0
      },
      "ILS": {
        "code": "ILS",
        "symbol": "₪",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": true,
        "decimalDigits": 2
      },
      "INR": {
        "code": "INR",
        "symbol": "₹",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "IQD": {
        "code": "IQD",
        "symbol": "د.ع.‏",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": true,
        "decimalDigits": 2
      },
      "IRR": {
        "code": "IRR",
        "symbol": "﷼",
        "thousandsSeparator": ",",
        "decimalSeparator": "/",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": true,
        "decimalDigits": 2
      },
      "ISK": {
        "code": "ISK",
        "symbol": "kr.",
        "thousandsSeparator": ".",
        "decimalSeparator": ",",
        "symbolOnLeft": false,
        "spaceBetweenAmountAndSymbol": true,
        "decimalDigits": 0
      },
      "JMD": {
        "code": "JMD",
        "symbol": "J$",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "JOD": {
        "code": "JOD",
        "symbol": "د.ا.‏",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": true,
        "decimalDigits": 3
      },
      "JPY": {
        "code": "JPY",
        "symbol": "¥",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 0
      },
      "KES": {
        "code": "KES",
        "symbol": "S",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "KGS": {
        "code": "KGS",
        "symbol": "сом",
        "thousandsSeparator": " ",
        "decimalSeparator": "-",
        "symbolOnLeft": false,
        "spaceBetweenAmountAndSymbol": true,
        "decimalDigits": 2
      },
      "KHR": {
        "code": "KHR",
        "symbol": "៛",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": false,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 0
      },
      "KMF": {
        "code": "KMF",
        "symbol": "CF",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": false,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "KPW": {
        "code": "KPW",
        "symbol": "₩",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 0
      },
      "KRW": {
        "code": "KRW",
        "symbol": "₩",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 0
      },
      "KWD": {
        "code": "KWD",
        "symbol": "د.ك.‏",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": true,
        "decimalDigits": 3
      },
      "KYD": {
        "code": "KYD",
        "symbol": "$",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "KZT": {
        "code": "KZT",
        "symbol": "₸",
        "thousandsSeparator": " ",
        "decimalSeparator": "-",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "LAK": {
        "code": "LAK",
        "symbol": "₭",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": false,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 0
      },
      "LBP": {
        "code": "LBP",
        "symbol": "ل.ل.‏",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": true,
        "decimalDigits": 2
      },
      "LKR": {
        "code": "LKR",
        "symbol": "₨",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": true,
        "decimalDigits": 0
      },
      "LRD": {
        "code": "LRD",
        "symbol": "$",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "LSL": {
        "code": "LSL",
        "symbol": "M",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": false,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "LYD": {
        "code": "LYD",
        "symbol": "د.ل.‏",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 3
      },
      "MAD": {
        "code": "MAD",
        "symbol": "د.م.‏",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": true,
        "decimalDigits": 2
      },
      "MDL": {
        "code": "MDL",
        "symbol": "lei",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": false,
        "spaceBetweenAmountAndSymbol": true,
        "decimalDigits": 2
      },
      "MGA": {
        "code": "MGA",
        "symbol": "Ar",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 0
      },
      "MKD": {
        "code": "MKD",
        "symbol": "ден.",
        "thousandsSeparator": ".",
        "decimalSeparator": ",",
        "symbolOnLeft": false,
        "spaceBetweenAmountAndSymbol": true,
        "decimalDigits": 2
      },
      "MMK": {
        "code": "MMK",
        "symbol": "K",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "MNT": {
        "code": "MNT",
        "symbol": "₮",
        "thousandsSeparator": " ",
        "decimalSeparator": ",",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "MOP": {
        "code": "MOP",
        "symbol": "MOP$",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "MRO": {
        "code": "MRO",
        "symbol": "UM",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": false,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "MTL": {
        "code": "MTL",
        "symbol": "₤",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "MUR": {
        "code": "MUR",
        "symbol": "₨",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "MVR": {
        "code": "MVR",
        "symbol": "MVR",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": false,
        "spaceBetweenAmountAndSymbol": true,
        "decimalDigits": 1
      },
      "MWK": {
        "code": "MWK",
        "symbol": "MK",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "MXN": {
        "code": "MXN",
        "symbol": "$",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "MYR": {
        "code": "MYR",
        "symbol": "RM",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "MZN": {
        "code": "MZN",
        "symbol": "MT",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 0
      },
      "NAD": {
        "code": "NAD",
        "symbol": "$",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "NGN": {
        "code": "NGN",
        "symbol": "₦",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "NIO": {
        "code": "NIO",
        "symbol": "C$",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": true,
        "decimalDigits": 2
      },
      "NOK": {
        "code": "NOK",
        "symbol": "kr",
        "thousandsSeparator": " ",
        "decimalSeparator": ",",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": true,
        "decimalDigits": 2
      },
      "NPR": {
        "code": "NPR",
        "symbol": "₨",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "NZD": {
        "code": "NZD",
        "symbol": "$",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "OMR": {
        "code": "OMR",
        "symbol": "﷼",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": true,
        "decimalDigits": 3
      },
      "PAB": {
        "code": "PAB",
        "symbol": "B/.",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": true,
        "decimalDigits": 2
      },
      "PEN": {
        "code": "PEN",
        "symbol": "S/.",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": true,
        "decimalDigits": 2
      },
      "PGK": {
        "code": "PGK",
        "symbol": "K",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "PHP": {
        "code": "PHP",
        "symbol": "₱",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "PKR": {
        "code": "PKR",
        "symbol": "₨",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "PLN": {
        "code": "PLN",
        "symbol": "zł",
        "thousandsSeparator": " ",
        "decimalSeparator": ",",
        "symbolOnLeft": false,
        "spaceBetweenAmountAndSymbol": true,
        "decimalDigits": 2
      },
      "PYG": {
        "code": "PYG",
        "symbol": "₲",
        "thousandsSeparator": ".",
        "decimalSeparator": ",",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": true,
        "decimalDigits": 2
      },
      "QAR": {
        "code": "QAR",
        "symbol": "﷼",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": true,
        "decimalDigits": 2
      },
      "RON": {
        "code": "RON",
        "symbol": "lei",
        "thousandsSeparator": ".",
        "decimalSeparator": ",",
        "symbolOnLeft": false,
        "spaceBetweenAmountAndSymbol": true,
        "decimalDigits": 2
      },
      "RSD": {
        "code": "RSD",
        "symbol": "Дин.",
        "thousandsSeparator": ".",
        "decimalSeparator": ",",
        "symbolOnLeft": false,
        "spaceBetweenAmountAndSymbol": true,
        "decimalDigits": 2
      },
      "RUB": {
        "code": "RUB",
        "symbol": "₽",
        "thousandsSeparator": " ",
        "decimalSeparator": ",",
        "symbolOnLeft": false,
        "spaceBetweenAmountAndSymbol": true,
        "decimalDigits": 2
      },
      "RWF": {
        "code": "RWF",
        "symbol": "RWF",
        "thousandsSeparator": " ",
        "decimalSeparator": ",",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": true,
        "decimalDigits": 2
      },
      "SAR": {
        "code": "SAR",
        "symbol": "﷼",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": true,
        "decimalDigits": 2
      },
      "SBD": {
        "code": "SBD",
        "symbol": "$",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "SCR": {
        "code": "SCR",
        "symbol": "₨",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "SDD": {
        "code": "SDD",
        "symbol": "LSd",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": false,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "SDG": {
        "code": "SDG",
        "symbol": "£‏",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "SEK": {
        "code": "SEK",
        "symbol": "kr",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": false,
        "spaceBetweenAmountAndSymbol": true,
        "decimalDigits": 2
      },
      "SGD": {
        "code": "SGD",
        "symbol": "$",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "SHP": {
        "code": "SHP",
        "symbol": "£",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "SLL": {
        "code": "SLL",
        "symbol": "Le",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "SOS": {
        "code": "SOS",
        "symbol": "S",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "SRD": {
        "code": "SRD",
        "symbol": "$",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "STD": {
        "code": "STD",
        "symbol": "Db",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "SVC": {
        "code": "SVC",
        "symbol": "₡",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "SYP": {
        "code": "SYP",
        "symbol": "£",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": true,
        "decimalDigits": 2
      },
      "SZL": {
        "code": "SZL",
        "symbol": "E",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "THB": {
        "code": "THB",
        "symbol": "฿",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "TJS": {
        "code": "TJS",
        "symbol": "TJS",
        "thousandsSeparator": " ",
        "decimalSeparator": ";",
        "symbolOnLeft": false,
        "spaceBetweenAmountAndSymbol": true,
        "decimalDigits": 2
      },
      "TMT": {
        "code": "TMT",
        "symbol": "m",
        "thousandsSeparator": " ",
        "decimalSeparator": ",",
        "symbolOnLeft": false,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 0
      },
      "TND": {
        "code": "TND",
        "symbol": "د.ت.‏",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": true,
        "decimalDigits": 3
      },
      "TOP": {
        "code": "TOP",
        "symbol": "T$",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "TRY": {
        "code": "TRY",
        "symbol": "TL",
        "thousandsSeparator": ".",
        "decimalSeparator": ",",
        "symbolOnLeft": false,
        "spaceBetweenAmountAndSymbol": true,
        "decimalDigits": 2
      },
      "TTD": {
        "code": "TTD",
        "symbol": "TT$",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "TVD": {
        "code": "TVD",
        "symbol": "$",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "TWD": {
        "code": "TWD",
        "symbol": "NT$",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "TZS": {
        "code": "TZS",
        "symbol": "TSh",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "UAH": {
        "code": "UAH",
        "symbol": "₴",
        "thousandsSeparator": " ",
        "decimalSeparator": ",",
        "symbolOnLeft": false,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "UGX": {
        "code": "UGX",
        "symbol": "USh",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "USD": {
        "code": "USD",
        "symbol": "$",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "UYU": {
        "code": "UYU",
        "symbol": "$U",
        "thousandsSeparator": ".",
        "decimalSeparator": ",",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": true,
        "decimalDigits": 2
      },
      "UZS": {
        "code": "UZS",
        "symbol": "сўм",
        "thousandsSeparator": " ",
        "decimalSeparator": ",",
        "symbolOnLeft": false,
        "spaceBetweenAmountAndSymbol": true,
        "decimalDigits": 2
      },
      "VEB": {
        "code": "VEB",
        "symbol": "Bs.",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "VEF": {
        "code": "VEF",
        "symbol": "Bs. F.",
        "thousandsSeparator": ".",
        "decimalSeparator": ",",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": true,
        "decimalDigits": 2
      },
      "VND": {
        "code": "VND",
        "symbol": "₫",
        "thousandsSeparator": ".",
        "decimalSeparator": ",",
        "symbolOnLeft": false,
        "spaceBetweenAmountAndSymbol": true,
        "decimalDigits": 1
      },
      "VUV": {
        "code": "VUV",
        "symbol": "VT",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": false,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 0
      },
      "WST": {
        "code": "WST",
        "symbol": "WS$",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "XAF": {
        "code": "XAF",
        "symbol": "F",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": false,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "XCD": {
        "code": "XCD",
        "symbol": "$",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "XOF": {
        "code": "XOF",
        "symbol": "F",
        "thousandsSeparator": " ",
        "decimalSeparator": ",",
        "symbolOnLeft": false,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "XPF": {
        "code": "XPF",
        "symbol": "F",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": false,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "YER": {
        "code": "YER",
        "symbol": "﷼",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": true,
        "decimalDigits": 2
      },
      "ZAR": {
        "code": "ZAR",
        "symbol": "R",
        "thousandsSeparator": " ",
        "decimalSeparator": ",",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "ZMW": {
        "code": "ZMW",
        "symbol": "ZK",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      },
      "WON": {
        "code": "WON",
        "symbol": "₩",
        "thousandsSeparator": ",",
        "decimalSeparator": ".",
        "symbolOnLeft": true,
        "spaceBetweenAmountAndSymbol": false,
        "decimalDigits": 2
      }
    }[code];
  }
  
  function isString(obj) {
    return !!(obj === '' || (obj && obj.charCodeAt && obj.substr));
  }
  
  function isObject(obj) {
    return obj && toString.call(obj) === '[object Object]';
  }
  