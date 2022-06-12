using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excelam.System;


/// <summary>
/// Currency code.
/// </summary>
public enum ExcelCellCurrencyCode
{
    Undefined,
    
    Unknown,

    /// <summary>
    /// general case
    /// </summary>
    Euro,

    /// <summary>
    /// United States dollar
    /// </summary>
    UnitedStatesDollar,

    AustralianDollar,

    CanadianDollar,

    /// <summary>
    /// Pound sterling
    /// United kingdom,..
    /// </summary>
    PoundSterling,

    ///// <summary>
    ///// ¥
    ///// japanese, yen
    ///// 
    ///// todo: pb meme symbole que china!
    ///// </summary>
    JapaneseYen,

    ChineseChina,

    YiChina,

    SingaporeDollar,

    /// <summary>
    /// Ukrainian hryvnia
    /// </summary>
    UkrainianHryvnia,

    /// <summary>
    /// 	Russian ruble
    /// </summary>
    RussianRuble,


    Bitcoin

    ///// <summary>
    ///// yuan (china)
    ///// renminbi , chinese
    ///// 
    ///// JP¥50 and CN¥50 when disambiguation is needed. 
    ///// </summary>
    //CurrencyChinese,

    ///// <summary>
    ///// South Korean
    ///// </summary>
    //CurrencyWon,

}

