using Excelam.System;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excelam;

/// <summary>
/// Decode from OpenXml number format to ExcelCellFormatValue high level objects.
/// </summary>
public class OxExcelCellFormatValueDecoder
{
	/// <summary>
	/// convert all ExcelNumberingFormat into ExcelCellFormatValue.
	/// </summary>
	/// <param name="listExcelNumberingFormat"></param>
	/// <returns></returns>
	public static void Decode(List<ExcelNumberingFormat> listExcelNumberingFormat)
	{
		listExcelNumberingFormat.ForEach(excelNumberingFormat =>
		{
			// clean
			if (excelNumberingFormat.FormatCode == null)
				excelNumberingFormat.FormatCode = string.Empty;

			ExcelCellFormatCode code;
			ExcelCellCurrencyCode currencyCode;
			DecodeNumberingFormat(excelNumberingFormat.Id, excelNumberingFormat.FormatCode, out code, out currencyCode);
			// set the decoded code
			excelNumberingFormat.Code = code;
			excelNumberingFormat.CurrencyCode = currencyCode;
		});  
	}


	public static void DecodeNumberingFormat(int numberFormatId, string format, out ExcelCellFormatCode code, out ExcelCellCurrencyCode countryCurrency)
	{
		countryCurrency = ExcelCellCurrencyCode.Undefined;
		// decode standard/default cases
		if (DecodeStandardCases(numberFormatId, out code))
			return;

		if(DecodeDateAndTimeCases(numberFormatId, format, out code))
			return;
		if (DecodeAccounting44Case(numberFormatId, format, out code, out countryCurrency))
			return;

		if (DecodeCurrencyCases(numberFormatId, format, out code, out countryCurrency))
			return;

		if (DecodeCurrencySpecialCase(numberFormatId, format, out code))
			return;

		// decode special math cases: fraction and percentage
		DecodeMathSpecialCases(numberFormatId, format, out code);
	}

	private static bool DecodeStandardCases(int numberFormatId, out ExcelCellFormatCode code)
	{
		code = ExcelCellFormatCode.Undefined;
		if (numberFormatId > 163)
			return false;

		if (numberFormatId == (int)ExcelCellBuiltInFormatCode.General)
		{
			code = ExcelCellFormatCode.General;
			return true;
		}

		if (numberFormatId == (int)ExcelCellBuiltInFormatCode.Text)
		{
			code = ExcelCellFormatCode.Text;
			return true;
		}

		if (numberFormatId == (int)ExcelCellBuiltInFormatCode.Number)
		{
			code = ExcelCellFormatCode.Number;
			return true;
		}

		if (numberFormatId == (int)ExcelCellBuiltInFormatCode.Decimal)
		{
			code = ExcelCellFormatCode.Decimal;
			return true;
		}

		if (numberFormatId == (int)ExcelCellBuiltInFormatCode.PercentageInt)
		{
			code = ExcelCellFormatCode.Percentage1;
			return true;
		}

		if (numberFormatId == (int)ExcelCellBuiltInFormatCode.Percentage2Dec)
		{
			code = ExcelCellFormatCode.Percentage2;
			return true;
		}

		if (numberFormatId == (int)ExcelCellBuiltInFormatCode.Scientific)
		{
			code = ExcelCellFormatCode.Scientific;
			return true;
		}

		if (numberFormatId == (int)ExcelCellBuiltInFormatCode.Fraction)
		{
			code = ExcelCellFormatCode.Fraction;
			return true;
		}

		if (numberFormatId == (int)ExcelCellBuiltInFormatCode.Fraction2Digit)
		{
			code = ExcelCellFormatCode.Fraction2Digit;
			return true;
		}

		if (numberFormatId == (int)ExcelCellBuiltInFormatCode.DateShort)
		{
			code = ExcelCellFormatCode.DateShort;
			return true;
		}

		// not a built-in case
		return false;
	}

	private static bool DecodeDateAndTimeCases(int numberFormatId, string formatCode, out ExcelCellFormatCode code)
	{
		code = ExcelCellFormatCode.Undefined;

		if (numberFormatId < 164)
			// its a built-in format, bye
			return false;


		// [$-F800]dddd\\,\\ mmmm\\ dd\\,\\ yyyy
		if (formatCode.StartsWith("[$-F800]"))
		{
			code = ExcelCellFormatCode.DateLarge;
			return true;
		}

		// [$-F400]h:mm:ss\\ AM/PM
		if (formatCode.StartsWith("[$-F400]"))
		{
			code = ExcelCellFormatCode.Time;
			return true;
		}

		return false;
	}

	/// <summary>
	/// Decode accounting case, numberFormatId=44.
	/// It's a special case, find the currency symbol.
	/// exp with Euro: : "_-* #,##0.00\\ \"€\"_-;\\-* #,##0.00\\ \"€\"_-;_-* \"-\"??\\ \"€\"_-;_-@_-"
	/// </summary>
	/// <param name="numberFormatId"></param>
	/// <param name="formatCode"></param>
	/// <param name="code"></param>
	/// <returns></returns>
	private static bool DecodeAccounting44Case(int numberFormatId, string formatCode, out ExcelCellFormatCode code, out ExcelCellCurrencyCode currencyCode)
	{
		code = ExcelCellFormatCode.Undefined;
		currencyCode = ExcelCellCurrencyCode.Undefined;

		// not the case
		if (numberFormatId != 44)
			return false;

		DecodeCurrencyCode(formatCode, out currencyCode);

		code = ExcelCellFormatCode.Accounting;
		return true;
	}

	/// <summary>
	/// Decode special case: "#,##0.00\\ \"€\""
	/// Doesn't contain tag [xxx] 
	/// </summary>
	/// <param name="numberFormatId"></param>
	/// <param name="excelCellFormat"></param>
	/// <param name="valueFormat"></param>
	/// <returns></returns>
	private static bool DecodeCurrencySpecialCase(int numberFormatId, string formatCode, out ExcelCellFormatCode code)
	{
		// TODO: undefined! rework it!
		code = ExcelCellFormatCode.Undefined;

		// doesn't contains [xxx]
		if (formatCode.Contains("[") || formatCode.Contains("]"))
			return false;

		ExcelCellCurrencyCode currencyCode;
		if (!DecodeCurrencyCode(formatCode, out currencyCode))
			return false;

		return false;
	}

	/// <summary>
	/// decode currency casees.
	/// based on ISO 4217.
	/// </summary>
	/// <param name="numberFormatId"></param>
	/// <param name="formatCode"></param>
	/// <param name="code"></param>
	/// <param name="countryCurrency"></param>
	/// <returns></returns>
	private static bool DecodeCurrencyCases(int numberFormatId, string formatCode, out ExcelCellFormatCode code, out ExcelCellCurrencyCode countryCurrency)
	{
		code = ExcelCellFormatCode.Undefined;
		countryCurrency = ExcelCellCurrencyCode.Undefined;

		if (numberFormatId < 164)
			// its a built-in format, bye
			return false;

		if (!DecodeCurrencyCode(formatCode, out countryCurrency))
			return false;

		code = ExcelCellFormatCode.Currency;
		return true;
	}

	/// <summary>
	/// decode special math cases: fraction and percentage.
	/// Exp:
	///   164/0.0%
	///   165/0.000%
	///   166/#" "?/2
	///
	/// </summary>
	/// <param name="numberFormatId"></param>
	/// <param name="formatCode"></param>
	/// <param name="code"></param>
	/// <returns></returns>
	private static bool DecodeMathSpecialCases(int numberFormatId, string formatCode, out ExcelCellFormatCode code)
	{
		code = ExcelCellFormatCode.Undefined;

		if (numberFormatId < 164)
			// its a built-in format, bye
			return false;

		if (formatCode.Contains("0.0%"))
		{
			code = ExcelCellFormatCode.PercentageOneDotOne;
			return true;
		}

		if (formatCode.Contains("0.000%"))
		{
			code = ExcelCellFormatCode.PercentageOneDotThree;
			return true;
		}

		if (formatCode.Contains("#\" \"?/2"))
		{
			code = ExcelCellFormatCode.FractionByTwo;
			return true;
		}
		return false;

	}

	/// <summary>
	/// Decode the currency code.
	/// </summary>
	/// <param name="formatCode"></param>
	/// <param name="currencyCode"></param>
	/// <returns></returns>
	private static bool DecodeCurrencyCode(string formatCode, out ExcelCellCurrencyCode currencyCode)
	{
		
		// euro
		if (formatCode.Contains("\"€"))
		{
			currencyCode = ExcelCellCurrencyCode.Euro;
			return true;
		}

		// dollar US
		if (formatCode.Contains("[$$-409]"))
		{
			currencyCode = ExcelCellCurrencyCode.UnitedStatesDollar;
			return true;
		}

		// [$$-C09]
		if (formatCode.Contains("[$$-C09]"))
		{
			currencyCode = ExcelCellCurrencyCode.AustralianDollar;
			return true;
		}

		if (formatCode.Contains("[$$-1009]"))
		{
			currencyCode = ExcelCellCurrencyCode.CanadianDollar;
			return true;
		}

		// [$£-809] pound, 
		if (formatCode.Contains("[$£-809]"))
		{
			currencyCode = ExcelCellCurrencyCode.PoundSterling;
			return true;
		}

		// #,##0.00\\ [$?-422]		Ukraine
		if (formatCode.Contains("-422]"))
		{
			currencyCode = ExcelCellCurrencyCode.UkrainianHryvnia;
			return true;
		}


		// [$¥-411]#,##0.00		Japonais
		if (formatCode.Contains("-411]"))
		{
			currencyCode = ExcelCellCurrencyCode.JapaneseYen;
			return true;
		}

		// #,##0.00\\ [$?-419] Russian
		if (formatCode.Contains("-419]"))
		{
			currencyCode = ExcelCellCurrencyCode.RussianRuble;
			return true;
		}

		// [$$-1004]#,##0.00 Chinese - Singapore
		if (formatCode.Contains("-1004]"))
		{
			currencyCode = ExcelCellCurrencyCode.SingaporeDollar;
			return true;
		}

		// [$¥-804]#,##0.00  Chinese - China
		if (formatCode.Contains("-804]"))
		{
			currencyCode = ExcelCellCurrencyCode.China;
			return true;
		}

		// [$¥-478]#,##0.00 Chinese - China, diff avec 804??
		if (formatCode.Contains("-478]"))
		{
			currencyCode = ExcelCellCurrencyCode.China;
			return true;
		}

		// [$₿]\\ #,##0.000000, bitcoin 
		if (formatCode.StartsWith("[$₿]"))
		{
			currencyCode = ExcelCellCurrencyCode.Bitcoin;
			return true;
		}

		if (formatCode.Contains("[$$-"))
		{
			// dollar, not managed
			currencyCode = ExcelCellCurrencyCode.Unknown;
			return true;
		}

		currencyCode = ExcelCellCurrencyCode.Undefined;
		return false;
	}
}
