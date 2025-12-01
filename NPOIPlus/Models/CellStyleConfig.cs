using NPOI.SS.UserModel;
using System;

namespace NPOIPlus.Models
{
	/// <summary>
	/// 單元格樣式配置
	/// </summary>
	public class CellStyleConfig
	{
		/// <summary>
		/// 樣式快取鍵名（相同鍵名會復用樣式）
		/// 如果為 null 或空字符串，樣式將不會被快取
		/// </summary>
		public string Key { get; set; }

		/// <summary>
		/// 樣式設置器（只在樣式不存在於快取時才執行）
		/// </summary>
		public Action<ICellStyle> StyleSetter { get; set; }

		/// <summary>
		/// 建構函式
		/// </summary>
		/// <param name="key">樣式快取鍵名</param>
		/// <param name="styleSetter">樣式設置器</param>
		public CellStyleConfig(string key, Action<ICellStyle> styleSetter)
		{
			Key = key;
			StyleSetter = styleSetter;
		}

		/// <summary>
		/// 解構方法（支援元組語法）
		/// </summary>
		public void Deconstruct(out string key, out Action<ICellStyle> styleSetter)
		{
			key = Key;
			styleSetter = StyleSetter;
		}
	}
}

