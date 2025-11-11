using AchieveItSystemModels;
using AchieveItSystemUI.View;
using System;
using System.Timers;
using System.Windows;
using Timer = System.Timers.Timer;

namespace AchieveItSystemUI
{
    public class CoinWindowViewModel : NotifyPropertyBase
    {
        private CoinModel _coin;
        private Random _random;
        private Timer _animationTimer;
        private int _animationCount;
        private string _windowTitle = "抛硬币";

        public CoinModel Coin
        {
            get => _coin;
            set => Set(ref _coin, value);
        }

        public string WindowTitle
        {
            get => _windowTitle;
            set => Set(ref _windowTitle, value);
        }

        public CommandBase FlipCoinCommand { get; }
        public CommandBase BackPortalCommand { get; set; }

        public CoinWindowViewModel()
        {
            Coin = new CoinModel();
            _random = new Random();

            _animationTimer = new Timer(100);
            _animationTimer.Elapsed += AnimationTimer_Elapsed;

            FlipCoinCommand = new CommandBase(FlipCoin);
            BackPortalCommand = new CommandBase(BackPortalFunc);
        }

        private void FlipCoin(object parameter)
        {
            if (Coin.IsFlipping) return;

            Coin.IsFlipping = true;
            _animationCount = 0;
            _animationTimer.Start();
        }

        private void AnimationTimer_Elapsed(object sender, ElapsedEventArgs e)
        {
            Application.Current.Dispatcher.Invoke(() =>
            {
                Coin.CurrentSide = _random.Next(2) == 0 ? CoinSide.Heads : CoinSide.Tails;
                _animationCount++;

                if (_animationCount >= 20)
                {
                    StopAnimation();
                }
            });
        }

        private void StopAnimation()
        {
            _animationTimer.Stop();
            Coin.IsFlipping = false;

            var finalSide = _random.Next(2) == 0 ? CoinSide.Heads : CoinSide.Tails;
            Coin.CurrentSide = finalSide;
        }

        public void BackPortalFunc(object o)
        {
            PortalWindow portal = new PortalWindow();
            (o as Window).Close();
            portal.Show();
        }
    }
}