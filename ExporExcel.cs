using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint.WebControls;
using System.IO;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Text.RegularExpressions;
using System.Web;
using System.Globalization;
using System.Drawing;
using System.Web.UI.HtmlControls;
using System.Data;

namespace SharePoint.Dev.Framework.Util
{
	public static class ExportarExcel
	{
		#region Properties

		static string[] meses = new string[] { "", "Tipo de Documento Vinculado", "Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez" };

		#endregion

		#region Métodos exportar para excel

		private static void CreateExcel(string strFileName, SPGridView gv)
		{
			using (StringWriter sw = new StringWriter())
			{
				using (HtmlTextWriter htw = new HtmlTextWriter(sw))
				{
					// Create a form to contain the grid
					Table table = new Table();

					// add the header row to the table
					if (gv.HeaderRow != null)
					{
					PrepareControlForExport(gv.HeaderRow);
					GridViewRow headerRow = gv.HeaderRow;
					headerRow.BackColor = Color.FromArgb(160, 160, 160);
					headerRow.ForeColor = Color.FromArgb(255, 255, 255);
					table.Rows.Add(gv.HeaderRow);
					}

					// add each of the data rows to the table
					foreach (GridViewRow row in gv.Rows)
					{
						PrepareControlForExport(row);

						if (row.RowIndex % 2 != 0)
						{
							row.BackColor = Color.FromArgb(238, 238, 238);
						}
						row.ForeColor = Color.FromArgb(115, 115, 115);
						table.Rows.Add(row);
					}

					// add the footer row to the table
					if (gv.FooterRow != null)
					{
						PrepareControlForExport(gv.FooterRow);
						GridViewRow FooterRow = gv.FooterRow;
						FooterRow.BackColor = Color.FromArgb(160, 160, 160);
						FooterRow.ForeColor = Color.FromArgb(255, 255, 255);
						FooterRow.Font.Bold = true;
						table.Rows.Add(gv.FooterRow);
					}

					table.RenderControl(htw);

					Regex cmdRegex = new Regex("<a(.*?)a>");
					string htmlText = cmdRegex.Replace(sw.ToString(), string.Empty);
					cmdRegex = new Regex("<img src=\"/_layouts/images/plus.gif(.*?)/>");
					htmlText = cmdRegex.Replace(htmlText, string.Empty);

					CultureInfo culture = CultureInfo.GetCultureInfo("pt-BR");

					Encoding encoding = Encoding.GetEncoding(culture.TextInfo.ANSICodePage);

					byte[] bytes = encoding.GetBytes(htmlText);
					HttpContext.Current.Response.Clear();
					HttpContext.Current.Response.AddHeader("content-disposition", string.Format("attachment; filename={0}", strFileName));
					HttpContext.Current.Response.ContentType = "application/vnd.ms-excel";
					HttpContext.Current.Response.Charset = System.Text.Encoding.GetEncoding(culture.TextInfo.ANSICodePage).BodyName;
					HttpContext.Current.Response.ContentEncoding = Encoding.UTF32;

					//render the htmlwriter into the response
					HttpContext.Current.Response.BinaryWrite(bytes);
					HttpContext.Current.Response.Flush();
					HttpContext.Current.Response.Close();
					HttpContext.Current.Response.Clear();
				}
			}
		}

		private static void PrepareControlForExport(Control control)
		{
			for (int i = 0; i < control.Controls.Count; i++)
			{

				Control current = control.Controls[i];

				if (current is LinkButton)
				{
					control.Controls.Remove(current);
					control.Controls.AddAt(i, new LiteralControl((current as LinkButton).Text));
				}
				else if (current is ImageButton)
				{
					control.Controls.Remove(current);
					control.Controls.AddAt(i, new LiteralControl((current as ImageButton).AlternateText));
				}
				else if (current is HyperLink)
				{
					control.Controls.Remove(current);
					control.Controls.AddAt(i, new LiteralControl((current as HyperLink).Text));
				}
				else if (current is DropDownList)
				{
					control.Controls.Remove(current);
					control.Controls.AddAt(i, new LiteralControl((current as DropDownList).SelectedItem.Text));
				}
				else if (current is CheckBox)
				{
					control.Controls.Remove(current);
					control.Controls.AddAt(i, new LiteralControl((current as CheckBox).Checked ? "True" : "False"));
				}
				else if (current is HyperLink)
				{
					control.Controls.Remove(current);
				}
				else if (current is HtmlImage)
				{
					control.Controls.Remove(current);
				}
				else if (current is System.Web.UI.WebControls.Image)
				{
					control.Controls.Remove(current);
				}

				if (current.HasControls())
				{
					PrepareControlForExport(current);
				}
			}
		}

		public static void ExportToExcel(string nameFile, DataTable dt)
		{
			SPGridView spTempGridview = new SPGridView();
			spTempGridview.AllowFiltering = false;
			spTempGridview.AllowSorting = false;
			spTempGridview.AllowPaging = false;
			spTempGridview.AutoGenerateColumns = true;
			spTempGridview.DataSource = dt;
			spTempGridview.DataBind();

			CreateExcel(nameFile + ".xls", spTempGridview);
		}

		public static void ExportToExcel(string nameFile, SPGridView sp)
		{
			SPGridView spReportProject2 = sp;
			spReportProject2.AllowFiltering = false;
			spReportProject2.AllowSorting = false;
			spReportProject2.AllowPaging = false;
			spReportProject2.AutoGenerateColumns = false;
			spReportProject2.DataBind();

			CreateExcel(nameFile + ".xls", spReportProject2);
		}

		public static void ExportToExcel(string nameFile, TreeView tv, string[] nomeColunas)
		{
			int contadorColunas = 0;
			DataTable dtPrincipal = new DataTable();

			foreach (TreeNode tnode in tv.Nodes)
			{
				MontaDataTable(tnode, contadorColunas, dtPrincipal, nomeColunas);
				contadorColunas = 0;
			}

			ExportToExcel(nameFile, dtPrincipal);
		}

		public static void MontaDataTable(TreeNode nodePai, int contadorColunas, DataTable dtPrincipal, string[] nomeColunas)
		{
			string nomeColuna = (nomeColunas.Count() > contadorColunas ? nomeColunas[contadorColunas] : string.Empty);

			DataRow dr = dtPrincipal.NewRow();
			if (dtPrincipal.Columns.Count <= contadorColunas)
				dtPrincipal.Columns.Add(nomeColuna);

			if (!nodePai.Text.Contains("<table"))
			{
				FormataValorColuna(nodePai.Text, dr, contadorColunas, dtPrincipal);
			}
			else
				FormataValorTabelaHTML(dtPrincipal, nodePai.Text, dr, contadorColunas);

			if (nodePai.ChildNodes.Count > 0)
			{
				contadorColunas++;
				foreach (TreeNode nodeFilho in nodePai.ChildNodes)
				{
					MontaDataTable(nodeFilho, contadorColunas, dtPrincipal, nomeColunas);
				}
			}
		}

		private static string FormataValorColunaTabela(string valor)
		{
			string ExpressaoTagsHTML = "<[^>]*>";
			string ExpressaoFarolTagsIMG = "<img src='[^'][^>]*>";

			if (Regex.Match(valor, ExpressaoFarolTagsIMG).Value.Contains("Red"))
				valor = Regex.Replace(valor, ExpressaoFarolTagsIMG, "[Vermelho]");
			if (Regex.Match(valor, ExpressaoFarolTagsIMG).Value.Contains("Yellow"))
				valor = Regex.Replace(valor, ExpressaoFarolTagsIMG, "[Amarelo]");
			if (Regex.Match(valor, ExpressaoFarolTagsIMG).Value.Contains("Green"))
				valor = Regex.Replace(valor, ExpressaoFarolTagsIMG, "[Verde]");

			var valores = System.Text.RegularExpressions.Regex.Replace(valor, ExpressaoTagsHTML, " ").Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

			if (valores.Count() == 2 && valores.GetValue(1).ToString().ToLower().Equals("bloqueado"))
			{
				valor = valores.GetValue(0) + " - " + valores.GetValue(1);
			}
			else if (valores.Count() == 3 && valores.GetValue(2).ToString().ToLower().Equals("bloqueado"))
			{
				valor = valores.GetValue(0) + " - " + valores.GetValue(1) + " " + valores.GetValue(2);
			}
			else
				valor = string.Join("", valores);

			return valor;
		}

		private static void FormataValorColuna(string valor, DataRow dr, int contadorColunas, DataTable dtPrincipal)
		{
			string ExpressaoTagsHTML = "<[^>]*>";
			string ExpressaoFarolTagsIMG = "<img src='[^'][^>]*>";

			if (Regex.Match(valor, ExpressaoFarolTagsIMG).Value.Contains("Red"))
				valor = Regex.Replace(valor, ExpressaoFarolTagsIMG, "[Vermelho]");
			if (Regex.Match(valor, ExpressaoFarolTagsIMG).Value.Contains("Yellow"))
				valor = Regex.Replace(valor, ExpressaoFarolTagsIMG, "[Amarelo]");
			if (Regex.Match(valor, ExpressaoFarolTagsIMG).Value.Contains("Green"))
				valor = Regex.Replace(valor, ExpressaoFarolTagsIMG, "[Verde]");

			var valores = System.Text.RegularExpressions.Regex.Replace(valor, ExpressaoTagsHTML, " ").Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

			if (valores.Count() == 2 && valores.GetValue(1).ToString().ToLower().Equals("bloqueado"))
			{
				valor = valores.GetValue(0) + " - " + valores.GetValue(1);
			}
			else if (valores.Count() == 3 && valores.GetValue(2).ToString().ToLower().Equals("bloqueado"))
			{
				valor = valores.GetValue(0) + " - " + valores.GetValue(1) + " " + valores.GetValue(2);
			}
			else
			{
				if (valores.Count() == 2 && (valores.GetValue(1).Equals("[Vermelho]") || valores.GetValue(1).Equals("[Amarelo]") || valores.GetValue(1).Equals("[Verde]")))
				{
					dr[contadorColunas] = valores.GetValue(0);
					if (!dtPrincipal.Columns.Contains("Farol"))
						dtPrincipal.Columns.Add("Farol");
					
					dr[contadorColunas + 1] = valores.GetValue(1);
					dtPrincipal.Rows.Add(dr);
				}
				else
				{
					dr[contadorColunas] = string.Join("", valores);
					dtPrincipal.Rows.Add(dr);
				}
			}

			
		}

		private static void FormataValorTabelaHTML(DataTable dtPrincipal, string textoHTML, DataRow dr, int quantidadeColunas)
		{
			DataTable table = ParseTable(textoHTML);
			int contadorMeses = 0;

			if (table.Columns.Count == 15)
			{
				table.Columns.Remove(table.Columns[2]);
			}
			else
				contadorMeses += 1;

			int quantidadeColunasTabela = quantidadeColunas;
			foreach (DataRow tr in table.Rows)
			{
				foreach (DataColumn td in tr.Table.Columns)
				{
					if (dtPrincipal.Columns.Count <= quantidadeColunasTabela)
						dtPrincipal.Columns.Add(meses[contadorMeses]);

					dr[quantidadeColunasTabela] = FormataValorColunaTabela(tr[td].ToString());

					if (contadorMeses <= (tr.Table.Columns.Count))
						contadorMeses++;
					
					quantidadeColunasTabela++;
				}
				
				dtPrincipal.Rows.Add(dr);
				dr = dtPrincipal.NewRow();
				quantidadeColunasTabela = quantidadeColunas;
			}
		}

		#endregion

		#region Conversao tabela html

		private const RegexOptions ExpressionOptions = RegexOptions.Singleline | RegexOptions.Multiline | RegexOptions.IgnoreCase;

		private const string CommentPattern = "<!--(.*?)-->";
		private const string TablePattern = "<table[^>]*>(.*?)</table>";
		private const string HeaderPattern = "<th[^>]*>(.*?)</th>";
		private const string RowPattern = "<tr[^>]*>(.*?)</tr>";
		private const string CellPattern = "<td[^>]*>(.*?)</td>";

		public static DataTable ParseTable(string tableHtml)
		{
			string tableHtmlWithoutComments = tableHtml;

			DataTable dataTable = new DataTable();

			MatchCollection rowMatches = Regex.Matches(
			tableHtmlWithoutComments,
			RowPattern,
			ExpressionOptions);

			dataTable.Columns.AddRange(tableHtmlWithoutComments.Contains("<th")
			? ParseColumns(tableHtml)
			: GenerateColumns(rowMatches));

			ParseRows(rowMatches, dataTable);

			return dataTable;
		}

		private static void ParseRows(MatchCollection rowMatches, DataTable dataTable)
		{
			foreach (Match rowMatch in rowMatches)
			{
				if (!rowMatch.Value.Contains("<th"))
				{
					DataRow dataRow = dataTable.NewRow();

					MatchCollection cellMatches = Regex.Matches(
					rowMatch.Value,
					CellPattern,
					ExpressionOptions);

					for (int columnIndex = 0; columnIndex < cellMatches.Count; columnIndex++)
					{
						dataRow[columnIndex] = cellMatches[columnIndex].Groups[1].ToString();
					}

					dataTable.Rows.Add(dataRow);
				}
			}
		}

		private static DataColumn[] ParseColumns(string tableHtml)
		{
			MatchCollection headerMatches = Regex.Matches(
			tableHtml,
			HeaderPattern,
			ExpressionOptions);

			return (from Match headerMatch in headerMatches
			select new DataColumn(headerMatch.Groups[1].ToString())).ToArray();
		}

		private static DataColumn[] GenerateColumns(MatchCollection rowMatches)
		{
			int columnCount = Regex.Matches(
			rowMatches[0].ToString(),
			CellPattern,
			ExpressionOptions).Count;

			return (from index in Enumerable.Range(0, columnCount)
			select new DataColumn("Column " + Convert.ToString(index))).ToArray();
		}

	#endregion
	}
}