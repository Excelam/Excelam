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

		if (DecodePercentageCases(numberFormatId, format, out valueBase))
			return;

		if (DecodeFractionCases(numberFormatId, format, out valueBase))
			return;

		if (DecodeScientificCases(numberFormatId, format, out valueBase))
			return;

		if (DecodeAccounting44Case(numberFormatId, format, out valueBase))
			return;

		if (DecodeCurrencySpecialCase(numberFormatId, format, out valueBase))
			return;

		DecodeCurrencyCases(numberFormatId, format, out valueBase);
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

		if (numberFormatId == (int)ExcelCellBuiltInFormatCode.General0)
		{
			valueBase = new ExcelCellFormatValueGeneral();
			return true;
		}

		if (numberFormatId == (int)ExcelCellBuiltInFormatCode.Number1)
		{
			valueBase = new ExcelCellFormatValueNumber();
			return true;
		}

		if (numberFormatId == (int)ExcelCellBuiltInFormatCode.Text49)
		{
			valueBase = new ExcelCellFormatValueText();
			return true;
		}

		// not a built-in case
		return false;
	}

	private static bool DecodeDecimalCases(int numberFormatId, string format, out ExcelCellFormatValueBase valueBase)
	{
		ExcelCellFormatValueDecimal formatValue;

		// built-in format value
		if (numberFormatId == (int)ExcelCellBuiltInFormatCode.Decimal2)
		{
			formatValue = new ExcelCellFormatValueDecimal();
			formatValue.Define(2, false, ExcelCellValueNegativeOption.Default);
			valueBase = formatValue;
			return true;
		}

		// built-in format value
		if (numberFormatId == (int)ExcelCellBuiltInFormatCode.Decimal4BlankThousandSep)
		{
			formatValue = new ExcelCellFormatValueDecimal();
			//formatValue.SetDecimalCode(ExcelCellDecimalCode.Decimal4BlankThousandSep, 2);
			formatValue.Define(2, true, ExcelCellValueNegativeOption.Default);
			valueBase = formatValue;
			return true;
		}

		if (string.IsNullOrEmpty(format))
		{
			valueBase = null;
			return false;
		}

		if (format == "0.0")
		{
			formatValue = new ExcelCellFormatValueDecimal();
			//formatValue.SetDecimalCode(ExcelCellDecimalCode.DecimalN, 1);
			formatValue.Define(1, false, ExcelCellValueNegativeOption.Default);
			formatValue.StringFormat = format;
			valueBase = formatValue;
			return true;
		}

		if (format=="0.000")
		{
			formatValue = new ExcelCellFormatValueDecimal();
			//formatValue.SetDecimalCode(ExcelCellDecimalCode.DecimalN, 3);
			formatValue.Define(3, false, ExcelCellValueNegativeOption.Default);
			formatValue.StringFormat = format;
			valueBase = formatValue;
			return true;
		}

		// Decimal, 2 decimal, negative: red
		if (format == "0.00_ ;[Red]\\-0.00\\ ")
		{
			formatValue = new ExcelCellFormatValueDecimal();
			//formatValue.SetDecimalCode(ExcelCellDecimalCode.DecimalNegRed, 2);
			formatValue.Define(2, false, ExcelCellValueNegativeOption.RedWithSign);
			formatValue.StringFormat = format;
			valueBase = formatValue;
			return true;
		}

		// Decimal, 2 decimal, negative: red, no sign. format: "0.00;[Red]0.00"
		if (format == "0.00;[Red]0.00")
		{
			formatValue = new ExcelCellFormatValueDecimal();
			//formatValue.SetDecimalCode(ExcelCellDecimalCode.DecimalNegRedNoSign, 2);
			formatValue.Define(2, false, ExcelCellValueNegativeOption.RedWithoutSign);
			formatValue.StringFormat = format;
			valueBase = formatValue;
			return true;
		}

		// not a built-in case
		valueBase = null;
		return false;
	}

	/// <summary>
	/// Decode date and/ro time cases.
	/// </summary>
	/// <param name="numberFormatId"></param>
	/// <param name="formatCode"></param>
	/// <param name="formatValueBase"></param>
	/// <returns></returns>
	private static bool DecodeDateTimeCases(int numberFormatId, string formatCode, out ExcelCellFormatValueBase? formatValueBase)
	{
		ExcelCellFormatValueDateTime formatValue;

		// built-in, 14 dateShort
		if (numberFormatId == (int)ExcelCellBuiltInFormatCode.DateShort14)
		{
			formatValue = new ExcelCellFormatValueDateTime();
			formatValue.DateTimeCode = ExcelCellDateTimeCode.DateShort14;
			formatValueBase = formatValue;
			return true;
		}

		// built-in, 21 = 'hh:mm:ss'
		if (numberFormatId == (int)ExcelCellBuiltInFormatCode.Time21_hh_mm_ss)
		{
			formatValue = new ExcelCellFormatValueDateTime();
			formatValue.DateTimeCode = ExcelCellDateTimeCode.Time21_hh_mm_ss;
			formatValueBase = formatValue;
			return true;
		}

		if (string.IsNullOrEmpty(formatCode))
		{
			formatValueBase = null;
			return false;
		}

		// "yyyy\\-mm\\-dd;@"
		if (formatCode.Equals("yyyy\\-mm\\-dd;@"))
		{
			formatValue = new ExcelCellFormatValueDateTime();
			formatValue.DateTimeCode = ExcelCellDateTimeCode.Date_yyyy_mm_dd;
			formatValueBase = formatValue;
			return true;
		}

		// "[$-F800]dddd\\,\\ mmmm\\ dd\\,\\ yyyy"
		if (formatCode.StartsWith("[$-F800]"))
		{
			formatValue = new ExcelCellFormatValueDateTime();
			formatValue.DateTimeCode = ExcelCellDateTimeCode.DateLarge;
			formatValueBase = formatValue;
			return true;
		}

		// "[$-F400]h:mm:ss\\ AM/PM"
		if (formatCode.StartsWith("[$-F400]"))
		{
			formatValue = new ExcelCellFormatValueDateTime();
			formatValue.DateTimeCode = ExcelCellDateTimeCode.Time;
			formatValueBase = formatValue;
			return true;
		}

		// "[$-409]mmmm\\ d\\,\\ yyyy;@" 
		if (formatCode.Equals("[$-409]mmmm\\ d\\,\\ yyyy;@"))
		{
			formatValue = new ExcelCellFormatValueDateTime();
			formatValue.DateTimeCode = ExcelCellDateTimeCode.DateLargeEnglishUS;
			formatValueBase = formatValue;
			return true;
		}

		// "[$-407]d\\.\\ mmmm\\ yyyy;@"
		if (formatCode.Equals("[$-407]d\\.\\ mmmm\\ yyyy;@"))
		{
			formatValue = new ExcelCellFormatValueDateTime();
			formatValue.DateTimeCode = ExcelCellDateTimeCode.DateLargeGermanGermany;
			formatValueBase = formatValue;
			return true;
		}
		// "[$-807]d\\.\\ mmmm\\ yyyy;@"
		if (formatCode.Equals("[$-807]d\\.\\ mmmm\\ yyyy;@"))
		{
			formatValue = new ExcelCellFormatValueDateTime();
			formatValue.DateTimeCode = ExcelCellDateTimeCode.DateLargeGermanSwitzerland;
			formatValueBase = formatValue;
			return true;
		}

		// dateTime others case: day and hour
		if (formatCode.Contains("d") && formatCode.Contains("h"))
		{
			formatValue = new ExcelCellFormatValueDateTime();
			formatValue.DateTimeCode = ExcelCellDateTimeCode.DateTimeOtherCases;
			formatValueBase = formatValue;
			return true;
		}

		// only date: (day and month) or (month and year)
		if ((formatCode.Contains("d") && formatCode.Contains("m")) || (formatCode.Contains("m") && formatCode.Contains("y")))
		{
			formatValue = new ExcelCellFormatValueDateTime();
			formatValue.DateTimeCode = ExcelCellDateTimeCode.DateOtherCases;
			formatValueBase = formatValue;
			return true;
		}

		// Time others case: hour or second
		if (formatCode.Contains("h") || formatCode.Contains("s"))
		{
			formatValue = new ExcelCellFormatValueDateTime();
			formatValue.DateTimeCode = ExcelCellDateTimeCode.TimeOtherCases;
			formatValueBase = formatValue;
			return true;
		}

		formatValueBase = null;
		return false;
	}

	private static bool DecodePercentageCases(int numberFormatId, string formatCode, out ExcelCellFormatValueBase? formatValueBase)
	{
		ExcelCellFormatValuePercentage formatValue;

		//  9 = '0%'
		if (numberFormatId == (int)ExcelCellBuiltInFormatCode.Percentage9Int)
		{
			formatValue = new ExcelCellFormatValuePercentage();
			formatValue.PercentageCode = ExcelCellPercentageCode.Percentage9Int;
			formatValue.NumberOfDecimal = 0;
			formatValueBase = formatValue;
			return true;
		}

		// 10 = '0.00%'
		if (numberFormatId == (int)ExcelCellBuiltInFormatCode.Percentage10Decimal)
		{
			formatValue = new ExcelCellFormatValuePercentage();
			formatValue.PercentageCode = ExcelCellPercentageCode.Percentage10Decimal2;
			formatValue.NumberOfDecimal = 2;
			formatValueBase = formatValue;
			return true;
		}

		if (string.IsNullOrEmpty(formatCode))
		{
			formatValueBase = null;
			return false;
		}

		if (formatCode.Contains("0.0%"))
		{
			formatValue = new ExcelCellFormatValuePercentage();
			formatValue.PercentageCode = ExcelCellPercentageCode.PercentageN;
			formatValue.NumberOfDecimal = 1;
			formatValueBase = formatValue;
			return true;
		}

		if (formatCode.Contains("0.000%"))
		{
			formatValue = new ExcelCellFormatValuePercentage();
			formatValue.PercentageCode = ExcelCellPercentageCode.PercentageN;
			formatValue.NumberOfDecimal = 3;
			formatValueBase = formatValue;
			return true;
		}

		if (formatCode.Contains("%"))
		{
			formatValue = new ExcelCellFormatValuePercentage();
			formatValue.PercentageCode = ExcelCellPercentageCode.PercentageOtherCases;
			formatValue.NumberOfDecimal = 3;
			formatValueBase = formatValue;
			return true;
		}

		// other cases
		// TODO:

		formatValueBase = null;
		return false;
	}

	private static bool DecodeFractionCases(int numberFormatId, string formatCode, out ExcelCellFormatValueBase? formatValueBase)
	{
		ExcelCellFormatValueFraction formatValue;

		//  12 = '# ?/?'
		if (numberFormatId == (int)ExcelCellBuiltInFormatCode.Fraction12)
		{
			formatValue = new ExcelCellFormatValueFraction();
			formatValue.FractionCode = ExcelCellFractionCode.Fraction12;
			formatValueBase = formatValue;
			return true;
		}

		// 13 = '# ??/??'
		if (numberFormatId == (int)ExcelCellBuiltInFormatCode.Fraction13)
		{
			formatValue = new ExcelCellFormatValueFraction();
			formatValue.FractionCode = ExcelCellFractionCode.Fraction13;
			formatValueBase = formatValue;
			return true;
		}

		if (string.IsNullOrEmpty(formatCode))
		{
			formatValueBase = null;
			return false;
		}

		if (formatCode.Contains("#\" \"?/2"))
		{
			formatValue = new ExcelCellFormatValueFraction();
			formatValue.FractionCode = ExcelCellFractionCode.FractionByTwo;
			formatValueBase = formatValue;
			return true;
		}

		formatValueBase = null;
		return false;
	}

	private static bool DecodeScientificCases(int numberFormatId, string formatCode, out ExcelCellFormatValueBase? formatValueBase)
	{
		ExcelCellFormatValueScientific formatValue;

		//  11 = '0.00E+00'
		if (numberFormatId == (int)ExcelCellBuiltInFormatCode.Scientific11)
		{
			formatValue = new ExcelCellFormatValueScientific();
			formatValue.ScientificCode = ExcelCellScientificCode.Scientific11;
			formatValueBase = formatValue;
			return true;
		}

		if (string.IsNullOrEmpty(formatCode))
		{
			formatValueBase = null;
			return false;
		}

		if (formatCode.Contains("E+") || formatCode.Contains("E-"))
		{
			formatValue = new ExcelCellFormatValueScientific();
			formatValue.ScientificCode = ExcelCellScientificCode.ScientificOtherCases;
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

		if (string.IsNullOrEmpty(formatCode))
			return false;

		// doesn't contains [xxx]
		if (formatCode.Contains("[") || formatCode.Contains("]"))
			return false;

		ExcelCellCurrencyCode currencyCode;
		if (DecodeCurrencyCode(formatCode, out currencyCode))
		{
			ExcelCellFormatValueCurrency formatValue = new ExcelCellFormatValueCurrency();
			formatValue.CurrencyCode = currencyCode;
			valueBase = formatValue;
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
		valueBase = formatValue;
		return true;
	}

	/// <summary>
	/// Decode the currency code.
	/// </summary>
	/// <param name="formatCode"></param>
	/// <param name="currencyCode"></param>
	/// <returns></returns>
	private static bool DecodeCurrencyCode(string formatCode, out ExcelCellCurrencyCode currencyCode)
	{
		if(string.IsNullOrEmpty(formatCode))
		{
			currencyCode = ExcelCellCurrencyCode.Undefined;
			return false;
		}

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
