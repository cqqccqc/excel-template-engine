using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using System.Text.RegularExpressions;
using Mice.Utility.Object;
using Ctrip.Common.Utility;
using System.Dynamic;
using System.ComponentModel;
using Mice.Utility.Log;
namespace Mice.Utility.Office
{
    /// <summary>
    /// step 1: prepare datasource
    ///
    /// TemplateAllContentVo data = new TemplateAllContentVo();
    /// and fill value into data.
    ///
    /// step 2: new templateengine instance
    ///
    /// ExcelTemplateEngine template = new ExcelTemplateEngine(path);
    ///
    /// step 3: do a render
    ///
    /// template.Render("main", data);
    ///
    /// </summary>
    public class ExcelTemplateEngine
    {
        //template file path
        private string filePath = string.Empty;
        private Regex loopText = new Regex("{loop:([a-zA-Z_0-9]+):([a-zA-Z_0-9.]+)}", RegexOptions.Compiled);
        private Regex varNameText = new Regex("{([a-zA-Z_0-9]+)}", RegexOptions.Compiled);
        // to match all {xxx}
        private Regex matchAllText = new Regex("{([a-zA-Z_0-9:.#]+)}", RegexOptions.Compiled);
        // to match a line to be deleted.
        private Regex toBeDelText = new Regex("{(#ToBeDeleted#)}", RegexOptions.Compiled);
        // to if condition
        private Regex ifText = new Regex("{if:([a-zA-Z_0-9]+):([a-zA-Z_0-9]+)}", RegexOptions.Compiled);
        // to include statement {include:templatename:varname}
        private Regex includeText = new Regex("{include:([a-zA-Z_0-9]+):([a-zA-Z_0-9]+)}", RegexOptions.Compiled);
        // pakcage
        private ExcelPackage package;
        // expandoobject data to hold datasource.
        private IDictionary<string, object> dataSource = new ExpandoObject();
        /// <summary>
        /// 加载Excel模板
        /// </summary>
        /// <param name="_filePath">模板路径</param>
        /// <param name="_templateType"></param>
        public ExcelTemplateEngine(string _filePath)
        {
            filePath = _filePath;
            byte[] file = File.ReadAllBytes(filePath);
            MemoryStream ms = new MemoryStream(file);
            package = new ExcelPackage(ms);
        }
        /// <summary>
        /// 保存Excel
        /// </summary>
        /// <param name="ms"></param>
        /// <returns></returns>
        public Stream WriteToStream(MemoryStream ms)
        {
            package.SaveAs(ms);
            ms.Position = 0;
            Stream s = new MemoryStream();
            ms.CopyTo(s);
            s.Seek(0, SeekOrigin.Begin);
            return s;
        }
        /// <summary>
        /// save as
        /// </summary>
        /// <param name="filePath"></param>
        public void SaveAs(string filePath)
        {
            FileInfo newFile = new FileInfo(filePath);
            package.SaveAs(newFile);
        }
        /// <summary>
        /// to render Subtemplte with null object data
        /// </summary>
        /// <param name="templateName"></param>
        /// <param name="propertyName"></param>
        /// <returns></returns>
        public Boolean RenderSubTemplateWithEmptyData(string templateName, string propertyName)
        {
            // backup subtemplate and render it in a new sheet
            var wsSubTemplate = package.Workbook.Worksheets[templateName];
            wsSubTemplate.Hidden = OfficeOpenXml.eWorkSheetHidden.Hidden;//hidden it
            var wsSubTemplateRender = package.Workbook.Worksheets.Copy(templateName, propertyName);
            // to get a range for render.
            var rowCnt = wsSubTemplateRender.Dimension.End.Row;
            var colCnt = wsSubTemplateRender.Dimension.End.Column;
            for (var rowIndex = 1; rowIndex <= rowCnt; rowIndex++)
            {
                for (var colIndex = 1; colIndex <= colCnt; colIndex++)
                {
                    //do a var replacement
                    string cellValue = CommonFunc.ConvertObjectToString(wsSubTemplateRender.Cells[rowIndex, colIndex].Value);
                    Match cellMatch = matchAllText.Match(cellValue);
                    if (cellMatch.Success)
                    {
                        wsSubTemplateRender.Cells[rowIndex, colIndex].Value = null;
                    }
                }
            }
            return true;
        }
        /// <summary>
        /// Render var syntax {varname} in the cell
        /// </summary>
        /// <param name="cellObject">cell object</param>
        /// <param name="propertyName">data used for render</param>
        /// <returns></returns>
        public Boolean RenderCellWithVar(ExcelRange cellObject, string propertyName)
        {
            //do a var replacement
            #region
            string cellValue = CommonFunc.ConvertObjectToString(cellObject.Value);
            Match varMatch = varNameText.Match(cellValue);
            if (varMatch.Success)
            {
                string varName = varMatch.Groups[1].Value;
                try
                {
                    cellObject.Value = dataSource[propertyName].GetType().GetProperty(varName).GetValue(dataSource[propertyName], null);
                }
                catch (Exception e)
                {
                    LogHelper.Warn("RenderCellWithVar:cell render for varMatch failed", e.Message);
                }
            }
            #endregion
            return varMatch.Success;
        }
        /// <summary>
        /// Render Loop syntax {loop:subtemplatename:listvarname} in the cell during subtemplate render
        /// {loop:subtemplatename:propertyName.listvarname}
        /// </summary>
        /// <param name="cellObject">cell object</param>
        /// <param name="propertyName">data used for render</param>
        /// <returns></returns>
        public Boolean RenderCellWithLoopInSubTemplate(ExcelRange cellObject, string propertyName)
        {
            //do a var replacement
            string cellValue = CommonFunc.ConvertObjectToString(cellObject.Value);
            #region
            Match loopMatch = loopText.Match(cellValue);
            if (loopMatch.Success)
            {
                string subTemplateName = loopMatch.Groups[1].Value;
                string varsName = loopMatch.Groups[2].Value;
                // replace loop statement with a new one
                string newValue = "{loop:" + subTemplateName + ":" + propertyName + "." + varsName + "}";
                cellObject.Value = newValue;
            }
            #endregion
            return loopMatch.Success;
        }
        /// <summary>
        /// Render If syntax {if:conditionvarname:subtemplatename} in the cell
        /// or {if:true:subtemplatename},{if:false:subtemplatename}
        /// {include:subtemplatename:propertyName}
        /// </summary>
        /// <param name="cellObject">cell object</param>
        /// <param name="propertyName">data used for render</param>
        /// <returns></returns>
        public Boolean RenderCellWithIf(ExcelRange cellObject, string propertyName)
        {
            string cellValue = CommonFunc.ConvertObjectToString(cellObject.Value);
            object dataRender = dataSource[propertyName];
            #region
            // if statement
            // do a if parameter modification : to add a property name in the loop statement
            Match ifMatch = ifText.Match(cellValue);
            if (ifMatch.Success)
            {
                string conditionVarName = ifMatch.Groups[1].Value;
                string subTemplateName = ifMatch.Groups[2].Value;
                try
                {
                    Boolean evaluationResult = false;
                    if (conditionVarName == "true")
                        evaluationResult = true;
                    else if (conditionVarName == "false")
                        evaluationResult = false;
                    else
                    {
                        evaluationResult = evaluationResult = (bool)dataRender.GetType().GetProperty(conditionVarName).GetValue(dataRender, null);
                    }
                    string newValue = evaluationResult ? "{include:" + subTemplateName + ":" + propertyName + "}" : "{#ToBeDeleted#}";
                    cellObject.Value = newValue;
                }
                catch (Exception e)
                {
                    LogHelper.Warn("Render:cell render for ifMatch failed", e.Message);
                    //continue; // don't do replacement and return orignal value.
                }
            }
            #endregion
            return ifMatch.Success;
        }
        /// <summary>
        /// Render internal include syntax {include:templatename:varname}  in the cell
        /// </summary>
        /// <param name="cellObject">cell object </param>
        /// <param name="propertyName">data used for render</param>
        /// <param name="mainTemplateName">main worksheet name for render</param>
        /// <returns></returns>
        public Boolean RenderCellWithInclude(ExcelRange cellObject, string propertyName, string mainTemplateName)
        {
            string cellValue = CommonFunc.ConvertObjectToString(cellObject.Value);
            //object dataRender = dataSource[propertyName];
            // if statement
            // do a if parameter modification : to add a property name in the loop statement
            #region
            Match includeMatch = includeText.Match(cellValue);
            if (includeMatch.Success)
            {
                string subTemplateName = includeMatch.Groups[1].Value;
                string newPropertyName = includeMatch.Groups[2].Value;
                try
                {
                    RenderSubTemplate(subTemplateName, newPropertyName, mainTemplateName, new ExcelCellAddress(cellObject.Start.Row, cellObject.Start.Column), true);
                }
                catch (Exception e)
                {
                    LogHelper.Warn("Render:cell render for includeMatch failed", e.Message);
                    //continue;
                }
            }
            #endregion
            return includeMatch.Success;
        }
        /// <summary>
        /// Render loop syntax {loop:templatename:listvarname}  in the cell
        /// </summary>
        /// <param name="cellObject">cell object </param>
        /// <param name="propertyName">data used for render</param>
        /// <param name="mainTemplateName">main worksheet name for render</param>
        /// <returns></returns>
        public Boolean RenderCellWithLoop(ExcelRange cellObject, string propertyName, string mainTemplateName)
        {
            string cellValue = CommonFunc.ConvertObjectToString(cellObject.Value);
            object dataRender = dataSource[propertyName];
            var rowIndex = cellObject.Start.Row;
            var colIndex = cellObject.Start.Column;
            var wsMain = package.Workbook.Worksheets[mainTemplateName];
            #region
            Match loopMatch = loopText.Match(cellValue);
            if (loopMatch.Success)
            {
                // Finally, subtemplatename and varsname .
                string subTemplateName = loopMatch.Groups[1].Value; // subtemplate name
                string varsName = loopMatch.Groups[2].Value;   // variable name
                // to check varsName include a "." or not
                // if loop in the main template, there is no ".",
                // otherwise, loop be replace it with "dataName.varName", where dataName is a dymiac priority
                // which added during template render.
                string[] dataNames = varsName.Split('.');
                string dataName = string.Empty;
                dynamic list = null;
                // to get a list object. if failed, do nothing
                try
                {
                    switch (dataNames.Length)
                    {
                        case 2: // it is a loop from subtemplate
                            dataName = dataNames[0];
                            varsName = dataNames[1];
                            break;
                        case 1: // it is a loop in the main template
                            dataName="main";
                            varsName = dataNames[0];
                           break;
                        default:
                            return true;
                    }
                    list = dataSource[dataName].GetType().GetProperty(varsName).GetValue(dataSource[dataName], null);
                }
                catch (Exception e)
                {
                    return true;
                }
                int lengthOfList = list == null ? 0 : list.Count;
                Boolean emptyList = false;
                if (lengthOfList == 0)
                {
                    // Empty List and to fake there is one item in the list.
                    lengthOfList = 1;
                    emptyList = true;
                }
                for (var indexOfList = lengthOfList - 1; indexOfList >= 0; indexOfList--)
                {
                    // get a item and add this item in the data expodoobject.
                    object item = emptyList ? null : item = list[indexOfList];
                    string newPropertyName = subTemplateName + "_" + rowIndex.ToString() + "_" + indexOfList.ToString();
                    dataSource.Add(newPropertyName, item);
                    // do a subtemplate render in the independent sheet firstly.
                    RenderSubTemplate(subTemplateName, newPropertyName, mainTemplateName, new ExcelCellAddress(rowIndex, colIndex), indexOfList == 0);
                }
            }
            #endregion
            return loopMatch.Success;
        }
        /// <summary>
        /// Render SumbTemplate
        /// </summary>
        /// <param name="templateName"> subtemplate name</param>
        /// <param name="propertyName"> propertyName</param>
        /// <param name="dataRender"></param>
        /// <returns></returns>
        public Boolean RenderSubTemplate(string templateName, string propertyName, string mainTemplateName, ExcelCellAddress location, Boolean replaceCurrentRow)
        {
            // backup subtemplate and render it in a new sheet
            var wsSubTemplate = package.Workbook.Worksheets[templateName];
            var wsMain = package.Workbook.Worksheets[mainTemplateName];
            object dataRender = dataSource[propertyName];
            wsSubTemplate.Hidden = OfficeOpenXml.eWorkSheetHidden.Hidden;//hidden it
            var wsSubTemplateRender = package.Workbook.Worksheets.Copy(templateName, propertyName);
            // to get a range for render.
            var rowCnt = wsSubTemplateRender.Dimension.End.Row;
            var colCnt = wsSubTemplateRender.Dimension.End.Column;
            for (var rowIndex = 1; rowIndex <= rowCnt; rowIndex++)
            {
                for (var colIndex = 1; colIndex <= colCnt; colIndex++)
                {
                    // to check whether there is {xx}, if not, continue it
                    Match cellMatch = matchAllText.Match(CommonFunc.ConvertObjectToString(wsSubTemplateRender.Cells[rowIndex, colIndex].Value));
                    if (cellMatch.Success == false)
                        continue;
                    else if (dataRender == null) // in case it is a empty, just render it with empty.
                    {
                        wsSubTemplateRender.Cells[rowIndex, colIndex].Value = null;
                    }
                    //handle a variable syntax
                    if (RenderCellWithVar(wsSubTemplateRender.Cells[rowIndex, colIndex], propertyName))
                        continue;
                     //handle loop syntax : to add a property name in the loop statement
                    if (RenderCellWithLoopInSubTemplate(wsSubTemplateRender.Cells[rowIndex, colIndex], propertyName))
                        continue;
                    // handle if syntax : to convert if syntax to include syntax
                    if (RenderCellWithIf(wsSubTemplateRender.Cells[rowIndex, colIndex], propertyName))
                        continue;
                }
            }
            // copy it into main template.
            // recaculate
            rowCnt = wsSubTemplateRender.Dimension.End.Row;
            colCnt = wsSubTemplateRender.Dimension.End.Column;
            // if the last  in the loop, only rowCnt-1 row should be inserted. another row is "loop" row.
            if (replaceCurrentRow)
            {
                if (rowCnt - 1 > 0) //need to insert new rows
                {
                    wsMain.InsertRow(location.Row + 1, rowCnt - 1, location.Row);
                }
                wsSubTemplateRender.Cells[1, 1, rowCnt, colCnt].Copy(wsMain.Cells[location.Row, location.Column]);
            }
            else
            {
                wsMain.InsertRow(location.Row + 1, rowCnt, location.Row);
                wsSubTemplateRender.Cells[1, 1, rowCnt, colCnt].Copy(wsMain.Cells[location.Row + 1, location.Column]);
            }
            // remove this subtemplate render result
            package.Workbook.Worksheets.Delete(wsSubTemplateRender);
            return true;
        }
        /// <summary>
        /// Render Engine
        /// </summary>
        /// <param name="mainTemplateName">
        /// excel sheet name is required to be rendered.
        /// </param>
        /// <param name="dataRender">
        /// dataSource is used during rendered.
        /// </param>
        /// <returns>
        /// TURE.
        /// </returns>
        public Boolean Render(string mainTemplateName, object dataRender)
        {
            // backup template
            var wsMain = package.Workbook.Worksheets[mainTemplateName];
            var wsMainTemplate = package.Workbook.Worksheets.Copy(mainTemplateName, mainTemplateName + "bakcup");
            wsMainTemplate.Hidden = OfficeOpenXml.eWorkSheetHidden.Hidden;//hidden it
            //呈现excel之前 清空数据源
            dataSource.Clear();
            dataSource.Add("main", dataRender);
            // do a render from one line to another line
            //rowEnd and colEnd will be changed after one row is rendered. it will be recaculated dynamic later
            var rowEnd = wsMain.Dimension.End.Row;
            var colEnd = wsMain.Dimension.End.Column;
            for (var rowIndex = 1; rowIndex <= rowEnd; rowIndex++)
            {
                bool isRowToBeRenderAgain = false;
                for (var colIndex = 1; colIndex <= colEnd; colIndex++)
                {
                    // to check whether there is {xx}, if not, continue it
                    Match cellMatch = matchAllText.Match(CommonFunc.ConvertObjectToString(wsMain.Cells[rowIndex, colIndex].Value));
                    if (cellMatch.Success == false) continue;
                    // handle var syntax
                    if (RenderCellWithVar(wsMain.Cells[rowIndex, colIndex], "main"))
                        continue;
                    // handle loop syntax
                    if (RenderCellWithLoop(wsMain.Cells[rowIndex, colIndex], "main", mainTemplateName))
                    {
                        isRowToBeRenderAgain = true;
                        continue;
                    }
                    // handle toBeDel syntax
                    #region
                    Match toBeDelMatch = toBeDelText.Match(CommonFunc.ConvertObjectToString(wsMain.Cells[rowIndex, colIndex].Value));
                    if (toBeDelMatch.Success)
                    {
                        wsMain.DeleteRow(rowIndex, 1);
                        isRowToBeRenderAgain = true;
                        continue;
                    }
                    #endregion
                    // handle if syntax
                    if (RenderCellWithIf(wsMain.Cells[rowIndex, colIndex], "main"))
                    {
                        isRowToBeRenderAgain = true;
                        continue;
                    }
                    // handle include syntax
                    if (RenderCellWithInclude(wsMain.Cells[rowIndex, colIndex], "main", mainTemplateName))
                    {
                        isRowToBeRenderAgain = true;
                        continue;
                    }
                }
                rowIndex = isRowToBeRenderAgain ? --rowIndex : rowIndex;
                //rowEnd and colEnd will be changed after one row is rendered. it will be recaculated dynamic
                rowEnd = wsMain.Dimension.End.Row;
                colEnd = wsMain.Dimension.End.Column;
            }
            return true;
        }
    }
}