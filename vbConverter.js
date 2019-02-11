var vbConverter = (function(){
	var getdate = function(oDt){
//		return (Object.prototype.toString.call(oDt)==="[object String]") ? new Date(oDt) : oDt;
		var oRslt;
		if (Object.prototype.toString.call(oDt)==="[object String]"){
			var arDt=oDt.split("/");
			var iDay=parseInt(arDt[0],10),iMonth=parseInt(arDt[1],10)-1,iYear=parseInt(arDt[2],10);
			iYear+= arDt[2].length==2 ? (iYear>50 ? 1900 : 2000) : 0;
			var oRslt = new Date(iYear,iMonth,iDay);
		}
		else oRslt = oDt;
		return oRslt;
	};
	return {
		isdate: function (sDt) {
			var arDt=sDt.split("/");
			var bRslt = false;
			if (arDt.length==3 && vbConverter.isnumeric(arDt[0]) && vbConverter.isnumeric(arDt[1]) && vbConverter.isnumeric(arDt[2]) && (arDt[2].length==2 || arDt[2].length==4)){
				var iDay=parseInt(arDt[0],10),iMonth=parseInt(arDt[1],10)-1,iYear=parseInt(arDt[2],10);
				iYear+= arDt[2].length==2 ? (iYear>50 ? 1900 : 2000) : 0;
				var oDt = new Date(iYear,iMonth,iDay);
				if (iDay==oDt.getDate() && iMonth==oDt.getMonth() && iYear==oDt.getFullYear()) bRslt=true;
			}
			return bRslt;
		},
		isnumeric: function(oNum){
			return !isNaN(oNum) && isFinite(oNum);
		},
		isarray: function(obj){
			return Object.prototype.toString.call(obj) === "[object Array]";
		},
		ubound: function(ar){
			return vbConverter.isarray(ar) ? ar.length-1 : null;
		},
		lbound: function(ar){
			return vbConverter.isarray(ar) ? 0 : null;
		},
		formatnumber: function (num,decimalNum,bolLeadingZero,bolParens,bolCommas)
		/**********************************************************************
			IN:
				NUM - the number to format
				decimalNum - the number of decimal places to format the number to
				bolLeadingZero - true / false - display a leading zero for
												numbers between -1 and 1
				bolParens - true / false - use parenthesis around negative numbers
				bolCommas - put commas as number separators.
		 
			RETVAL:
				The formatted number!
		 **********************************************************************/
		{ 
				if (isNaN(parseInt(num))) return "NaN";

			var tmpNum = num;
			var iSign = num < 0 ? -1 : 1;		// Get sign of number
			
			// Adjust number so only the specified number of numbers after
			// the decimal point are shown.
			tmpNum *= Math.pow(10,decimalNum);
			tmpNum = Math.round(Math.abs(tmpNum))
			tmpNum /= Math.pow(10,decimalNum);
			tmpNum *= iSign;					// Readjust for sign
			
			
			// Create a string object to do our formatting on
			var tmpNumStr = new String(tmpNum);

			// See if we need to strip out the leading zero or not.
			if (!bolLeadingZero && num < 1 && num > -1 && num != 0)
				if (num > 0)
					tmpNumStr = tmpNumStr.substring(1,tmpNumStr.length);
				else
					tmpNumStr = "-" + tmpNumStr.substring(2,tmpNumStr.length);
				
			// See if we need to put in the commas
			if (bolCommas && (num >= 1000 || num <= -1000)) {
				var iStart = tmpNumStr.indexOf(".");
				if (iStart < 0)
					iStart = tmpNumStr.length;

				iStart -= 3;
				while (iStart >= 1) {
					tmpNumStr = tmpNumStr.substring(0,iStart) + "," + tmpNumStr.substring(iStart,tmpNumStr.length)
					iStart -= 3;
				}		
			}

			// See if we need to use parenthesis
			if (bolParens && num < 0)
				tmpNumStr = "(" + tmpNumStr.substring(1,tmpNumStr.length) + ")";

			return tmpNumStr;		// Return our formatted string!
		},
		trim: function(sTxt){
			return sTxt.replace(/^\s*|\s*$/ig,"");
		},
		len: function(sTxt){
			return sTxt.length;
		},
		left: function(sTxt,iLen){
			return sTxt.substr(0,iLen);
		},
		right: function(sTxt,iLen){
			return sTxt.replace(vbConverter.left(sTxt,sTxt.length-iLen),"");
		},
		mid: function(sTxt,iStart,iLen){
			iStart-=1;
			if (!iLen)iLen=sTxt.length-iStart;
			return sTxt.substr(iStart,iLen);
		},
		instr: function(sTxt,sSearchFor){
			return sTxt.indexOf(sSearchFor)==-1 ? 0 : sTxt.indexOf(sSearchFor)+1;
		},
		day: function(oDt){
			oDt=getdate(oDt);
			return oDt.getDate();
		},
		month: function(oDt){
			oDt=getdate(oDt);
			return oDt.getMonth()+1;
		},
		year: function(oDt){
			oDt=getdate(oDt);
			return oDt.getFullYear();
		},
		cint:function (sTxt){
			return parseInt(sTxt,10);
		},
		cstr:function (sTxt){
			return sTxt.toString();
		},
		cdate:function (sTxt){
			var arDt=sTxt.split("/"),oDt=null;
			if (arDt.length==3 && vbConverter.isnumeric(arDt[0]) && vbConverter.isnumeric(arDt[1]) && vbConverter.isnumeric(arDt[2]) && (arDt[2].length==2 || arDt[2].length==4)){
				var iDay=parseInt(arDt[0],10),iMonth=parseInt(arDt[1],10)-1,iYear=parseInt(arDt[2],10);
				iYear+= arDt[2].length==2 ? (iYear>50 ? 1900 : 2000) : 0;
				oDt = new Date(iYear,iMonth,iDay);
			}
			return oDt;
		},
		exp:function(iPow){
			return (Math.pow(Math.E,iPow));
		},
		clng:function(sTxt){
			return parseFloat(sTxt);
		},
		round:function (sNum,iNum){
			if (!iNum)iNum=0;
			return Math.round(parseFloat(sNum)*Math.pow(10,iNum))/Math.pow(10,iNum);
		},
		chr:function(iCode){
			return String.fromCharCode(iCode);
		},
		replace:function (sSrc,sWhat,sWith){
			var jj=0;
			while (sSrc.indexOf(sWhat)>-1 && jj<10000){
				sSrc=sSrc.replace(sWhat,sWith);
				jj++;
			}
			return sSrc;
		},
		datediff:function (sInterval,sDt1,sDt2){
			var oDt1=getdate(sDt1),oDt2=getdate(sDt2),iTime=oDt2.getTime()-oDt1.getTime(),iRslt=-1;
			switch (sInterval){
				case "d":
					iRslt=iTime/1000/3600/24;
					break;
				case "m":
					iRslt=iTime/1000/3600/24/30;
					break;
				case "yyyy":
					iRslt=iTime/1000/3600/24/365;
					break;
			}
			return Math.round(iRslt);
		},
		dateadd:function(sInterval,iCount,sDt){
			var oDt=getdate(sDt);
			switch (sInterval){
				case "d":
					oDt.setDate(oDt.getDate()+iCount);
					break;
				case "m":
					oDt.setMonth(oDt.getMonth()+iCount);
					break;
				case "yyyy":
					oDt.setFullYear(oDt.getFullYear()+iCount);
					break;
			}
			return oDt;
		}
	}
})();

