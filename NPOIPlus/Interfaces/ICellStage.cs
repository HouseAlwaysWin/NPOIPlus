using NPOIPlus.Helpers;

namespace NPOIPlus
{
	public interface ICellStage
	{
		ICellStage SetValue<T>(T value);
	}
}

