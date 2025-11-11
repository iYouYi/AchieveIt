using System;

namespace AchieveItSystemModels
{
    public enum CoinSide
    {
        Heads,   // 正面
        Tails    // 反面
    }

    public class CoinModel : NotifyPropertyBase
    {
        private CoinSide _currentSide;
        private bool _isFlipping;

        public CoinSide CurrentSide
        {
            get => _currentSide;
            set
            {
                if (Set(ref _currentSide, value))
                {
                    RaisePropertyChanged(nameof(SideText));
                    RaisePropertyChanged(nameof(SideEmoji));
                }
            }
        }

        public bool IsFlipping
        {
            get => _isFlipping;
            set => Set(ref _isFlipping, value);
        }

        public string SideText => CurrentSide == CoinSide.Heads ? "正面" : "反面";
        public string SideEmoji => CurrentSide == CoinSide.Heads ? "😊" : "😎";

        public CoinModel()
        {
            CurrentSide = CoinSide.Heads; // 默认显示正面
            IsFlipping = false;
        }
    }
}