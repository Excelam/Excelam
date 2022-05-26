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

			ExcelCellFormatValueBase formatValue;
			DecodeNumberingFormat(excelNumberingFormat.Id, excelNumberingFormat.FormatCode, out formatValue);
			// set the decoded code
			excelNumberingFormat.ValueBase= formatValue;
		});  
	}


	public static void DecodeNumberingFormat(int numberFormatId, string format, out ExcelCellFormatValueBase valueBase)
	{
		// decode basic cases: general, number and text
		if (DecodeBasicCases(numberFormatId, out valueBase))
			return;

		// decode decimal cases
		if (DecodeDecimalCases(numberFormatId, format, out valueBase))
			return;

		if (DecodeDateTimeCases(numberFormatId, format, out valueBase))
			return;

		if (DecodeAccounting44Case(numberFormatId, format, out valueBase))
			return;

		if (DecodeCurrencySpecialCase(numberFormatId, format, out valueBase))
			return;

		if (DecodeCurrencyCases(numberFormatId, format, out valueBase))
			return;


		// decode special math cases: fraction and percentage
		DecodeMathSpecialCases(numberFormatId, format, out valueBase);
	}

	/// <summary>
	/// Decode very basic cases: general, number and text.
	/// </summary>
	/// <param name="numberFormatId"></param>
	/// <param name="code"></param>
	/// <returns></returns>
	private static bool DecodeBasicCases(int numberFormatId, out ExcelCellFormatValueBase valueBase)
	{
		valueBase = null;
		if (numberFormatId > 163)
			return false;

		if (numberFormatId == (int)ExcelCellBuiltInFormatCode.General)
		{
			valueBase = new ExcelCellFormatValueGeneral();
			return true;
		}

		if (numberFormatId == (int)ExcelCellBuiltInFormatCode.Number)
		{
			valueBase = new ExcelCellFormatValueNumber();
			return true;
		}

		if (numberFormatId == (int)ExcelCellBuiltInFormatCode.Text)
		{
			valueBase = new ExcelCellFormatValueText();
			return true;
		}

		//if (numberFormatId == (int)ExcelCellBuiltInFormatCode.Number)
		//{
		//	code = new ExcelCellFormatStructCode();
		//	code.MainCode = ExcelCellFormatMainCode.Number;
		//	return true;
		//}

		//if (numberFormatId == (int)ExcelCellBuiltInFormatCode.Decimal)
		//{
		//	code = new ExcelCellFormatStructCode();
		//	code.MainCode = ExcelCellFormatMainCode.Decimal;
		//	return true;
		//}

		//if (numberFormatId == (int)ExcelCellBuiltInFormatCode.PercentageInt)
		//{
		//	code = new ExcelCellFormatStructCode();
		//	code.MainCode = ExcelCellFormatMainCode.Percentage1;
		//	return true;
		//}

		//if (numberFormatId == (int)ExcelCellBuiltInFormatCode.Percentage2Dec)
		//{
		//	code = new ExcelCellFormatStructCode();
		//	code.MainCode = ExcelCellFormatMainCode.Percentage2;
		//	return true;
		//}

		//if (numberFormatId == (int)ExcelCellBuiltInFormatCode.Scientific)
		//{
		//	code = new ExcelCellFormatStructCode();
		//	code.MainCode = ExcelCellFormatMainCode.Scientific;
		//	return true;
		//}

		//if (numberFormatId == (int)ExcelCellBuiltInFormatCode.Fraction)
		//{
		//	code = new ExcelCellFormatStructCode();
		//	code.MainCode = ExcelCellFormatMainCode.Fraction;
		//	return true;
		//}

		//if (numberFormatId == (int)ExcelCellBuiltInFormatCode.Fraction2Digit)
		//{
		//	code = new ExcelCellFormatStructCode();
		//	code.MainCode = ExcelCellFormatMainCode.Fraction2Digit;
		//	return true;
		//}

		//if (numberFormatId == (int)ExcelCellBuiltInFormatCode.DateShort)
		//{
		//	code = ExcelCellFormatCode.DateShort;
		//	return true;
		//}

		// not a built-in case
		return false;
	}

	private static bool DecodeDecimalCases(int numberFormatId, string format, out ExcelCellFormatValueBase valueBase)
	{
		ExcelCellFormatValueDecimal formatValue;

		if (numberFormatId == (int)ExcelCellBuiltInFormatCode.Decimal)
		{
			formatValue = new ExcelCellFormatValueDecimal();
			formatValue.NumberOfDecimal = 2;
			valueBase = formatValue;
			return true;
		}

		if (format == null)
		{
			valueBase = null;
			return false;
		}

		if (format == "0.0")
		{
			formatValue = new ExcelCellFormatValueDecimal();
			formatValue.NumberOfDecimal = 1;
			formatValue.StringFormat = format;
			valueBase = formatValue;
			return true;
		}

		if (format=="0.000")
		{
			formatValue = new ExcelCellFormatValueDecimal();
			formatValue.NumberOfDecimal = 3;
			formatValue.StringFormat = format;
			valueBase = formatValue;
			return true;

		}

		
		// Decimal, 2 decimal, negative: red, no sign. format: "0.00;[Red]0.00"
		// Decimal, 2 decimal, negative: red, format: "0.00_ ;[Red]\\-0.00\\ "

		// not a built-in case
		valueBase = null;
		return false;
	}

	private static bool DecodeDateTimeCases(int numberFormatId, string formatCode, out ExcelCellFormatValueBase? formatValueBase)
	{
		ExcelCellFormatValueDateTime formatValue;


		if (numberFormatId == (int)ExcelCellBuiltInFormatCode.DateShort)
		{
			formatValue = new ExcelCellFormatValueDateTime();
			formatValue.DateTimeCode = ExcelCellDateTimeCode.DateShort;
			formatValueBase = formatValue;
			return true;
		}


		// [$-F800]dddd\\,\\ mmmm\\ dd\\,\\ yyyy
		if (formatCode.StartsWith("[$-F800]"))
		{
			formatValue = new ExcelCellFormatValueDateTime();
			formatValue.DateTimeCode = ExcelCellDateTimeCode.DateLarge;
			formatValueBase = formatValue;
			return true;
		}

		// [$-F400]h:mm:ss\\ AM/PM
		if (formatCode.StartsWith("[$-F400]"))
		{
			formatValue = new ExcelCellFormatValueDateTime();
			formatValue.DateTimeCode = ExcelCellDateTimeCode.Time;
			formatValueBase = formatValue;
			return true;
		}

		formatValueBase = null;
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
	private static bool DecodeAccounting44Case(int numberFormatId, string formatCode, out ExcelCellFormatValueBase valueBase)
	{
		valueBase = null;
		// not the case
		if (numberFormatId != 44)
			return false;

		ExcelCellFormatValueAccounting formatValue = new ExcelCellFormatValueAccounting();

		ExcelCellCurrencyCode currencyCode;
		DecodeCurrencyCode(formatCode, out currencyCode);

		formatValue.CurrencyCode = currencyCode;
		valueBase = formatValue;
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
	private static bool DecodeCurrencySpecialCase(int numberFormatId, string formatCode, out ExcelCellFormatValueBase valueBase)
	{
		valueBase = null;

		// doesn't contains [xxx]
		if (formatCode.Contains("[") || formatCode.Contains("]"))
			return false;

		ExcelCellCurrencyCode currencyCode;
		if (DecodeCurrencyCode(formatCode, out currencyCode))
		{
			ExcelCellFormatValueCurrency formatValue = new ExcelCellFormatValueCurrency();
			formatValue.CurrencyCode = currencyCode;
			return true;
		}

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
	private static bool DecodeCurrencyCases(int numberFormatId, string formatCode, out ExcelCellFormatValueBase valueBase)
	{
		valueBase = null;

		if (numberFormatId < 164)
			// its a built-in format, bye
			return false;

		ExcelCellCurrencyCode currencyCode;
		if (!DecodeCurrencyCode(formatCode, out currencyCode))
			return false;

		ExcelCellFormatValueCurrency formatValue = new ExcelCellFormatValueCurrency();
		formatValue.CurrencyCode = currencyCode;
		return true;
	}

	/// <summary>
	/// TODO: splitter en plusieurs: percentage, fraction, scientific.
	/// 
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
	private static bool DecodeMathSpecialCases(int numberFormatId, string formatCode, out ExcelCellFormatValueBase valueBase)
	{
		valueBase = null;

		if (numberFormatId < 164)
			// its a built-in format, bye
			return false;

		if (formatCode.Contains("0.0%"))
		{
			//ExcelCellFormatValuePercentage formatValue = new ExcelCellFormatValuePercentage();
			//code = new ExcelCellFormatStructCode();
			// todo:
			//code.MainCode = ExcelCellFormatMainCode.Percentage1;
			//code.PercentageCode = ExcelCellFormatMainCode.PercentageOneDotOne;
			return true;
		}

		if (formatCode.Contains("0.000%"))
		{
			//code = new ExcelCellFormatStructCode();
			// todo:
			//code.MainCode = ExcelCellFormatMainCode.Percentage1;
			//code.PercentageCode = ExcelCellFormatMainCode.PercentageOneDotThree;
			return true;
		}

		if (formatCode.Contains("#\" \"?/2"))
		{
			//code = new ExcelCellFormatStructCode();
			// todo:
			//code.MainCode = ExcelCellFormatMainCode.Percentage1;
			//code = ExcelCellFormatMainCode.FractionByTwo;
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
