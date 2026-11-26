/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
 * See https://github.com/JanKallman/EPPlus for details.
 *
 * Copyright (C) 2011  Jan Källman
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.

 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  
 * See the GNU Lesser General Public License for more details.
 *
 * The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 *
 * All code and executables are provided "as is" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * Code change notes:
 * 
 * Author          Change						                  Date
 * ******************************************************************************
 * Eyal Seagull    Conditional Formatting Adaption    2012-04-03
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using OfficeOpenXml.ConditionalFormatting;
using System.Drawing;

namespace OfficeOpenXml.ConditionalFormatting.Contracts
{
	/// <summary>
	/// IExcelConditionalFormattingDataBar
	/// </summary>
	public interface IExcelConditionalFormattingDataBarGroup
        : IExcelConditionalFormattingRule
	{
		#region Public Properties
        /// <summary>
        /// ShowValue
        /// </summary>
        bool ShowValue { get; set; }
        /// <summary>
        /// Databar Low Value
        /// </summary>
        ExcelConditionalFormattingIconDataBarValue LowValue { get;  }

        /// <summary>
        /// Databar High Value
        /// </summary>
        ExcelConditionalFormattingIconDataBarValue HighValue { get; }
        /// <summary>
        /// The color of the databar
        /// </summary>
        Color Color { get; set;}

        // orabooster-epplus (item B4 of issue #8, issue #46): the options of Excel 2010 and later. See ExcelConditionalFormattingDataBar.
        /// <summary>True if the rule has the part of Excel 2010 and later, where the options below are.</summary>
        bool HasExcel2010Options { get; }
        /// <summary>How the bar is filled: a gradient (the default) or the colour.</summary>
        eDataBarFillType FillType { get; set; }
        /// <summary>True if the bar has a border.</summary>
        bool ShowBorder { get; set; }
        /// <summary>The colour of the border, or null.</summary>
        Color? BorderColor { get; set; }
        /// <summary>The colour of the bars of negative values; null for the colour of the bar.</summary>
        Color? NegativeFillColor { get; set; }
        /// <summary>The colour of the border of the bars of negative values; null for the colour of the border of the bar.</summary>
        Color? NegativeBorderColor { get; set; }
        /// <summary>Where the axis between the bars of negative and of positive values is.</summary>
        eDataBarAxisPosition AxisPosition { get; set; }
        /// <summary>The colour of the axis, or null.</summary>
        Color? AxisColor { get; set; }
        /// <summary>The direction of the bar.</summary>
        eDataBarDirection Direction { get; set; }
        /// <summary>The length of the shortest bar in percent of the width of the cell, 0 to 100 (default 10).</summary>
        int MinLength { get; set; }
        /// <summary>The length of the longest bar in percent of the width of the cell, 0 to 100 (default 90).</summary>
        int MaxLength { get; set; }
        #endregion Public Properties
	}
}