#region License

// The MIT License (MIT)
//
// Copyright (c) 2021 Richard L King (TradeWright Software Systems)
// 
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
// 
// The above copyright notice and this permission notice shall be included in all
// copies or substantial portions of the Software.
// 
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
// SOFTWARE.

#endregion

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;

using AxTradingUI27;
using BarFormatters27;
using BarUtils27;
using ChartSkil27;
using ChartUtils27;
using ContractUtils27;
using HistDataUtils27;
using MarketDataUtils27;
using StudyUtils27;
using TickUtils27;
using TimeframeUtils27;
using TradingUI27;
using TWUtilities40;

namespace TradeWright.TradeBuild.Applications.Chart
{
    public partial class FChart : Form
    {
        private static _TWUtilities 
        TW;

        private static readonly _ChartSkil 
        ChartSkil = new ChartSkil();

        private static readonly _ContractUtils
        Contracts = new ContractUtils();

        private static readonly _TimeframeUtils
        TF = new TimeframeUtils();

        private static readonly _MarketDataUtils
        MD = new MarketDataUtils();

        private readonly StudyLibraryManager
        mStudyLibraryManager = new StudyLibraryManager();

        private ApiClientManager
        mClientManager;

        private _IMarketDataManager
        mMarketDataManager;

        private ConsoleHandler1
        mConsoleHandler;

        private IHistoricalDataStore
        mHistDataStore;

        private ChartTabPage
        mCurrentTabPage;


        internal FChart()
        {
            InitializeComponent();
            TW = new TWUtilities();

            mStudyLibraryManager.AddBuiltInStudyLibrary();

            tabControlExtra1.Click += (s, e) =>
            {
                if (mCurrentTabPage == (tabControlExtra1.SelectedTab as ChartTabPage))
                {
                    selectChart(mCurrentTabPage);
                }
            };

            tabControlExtra1.SelectedIndexChanged += (s, e) =>
            {
                if (mCurrentTabPage != null)
                {
                    mCurrentTabPage.Context.IsCurrent = false;
                }

                if (tabControlExtra1.SelectedIndex != -1)
                {
                    mCurrentTabPage = tabControlExtra1.SelectedTab as ChartTabPage;
                    selectChart(mCurrentTabPage);
                }
            };


            ChartTree.AfterSelect += (s, e) =>
            {
                if (e.Node is ChartTreeNode node)
                {
                    var context = node.Context;
                    if (context.TabPage == null)
                    {
                        AddChart(context);
                    }
                    else
                    {
                        tabControlExtra1.SelectedTab = node.Context.TabPage;
                    }
                }
            };

            void selectChart(ChartTabPage tabPage)
            {
                var chartContext = tabPage.Context;
                chartContext.IsCurrent = true;
                if (chartContext.MultiChart.CurrentIndex == 0 && chartContext.MultiChart.Count > 0)
                {
                    chartContext.MultiChart.SelectChart(chartContext.InitialChartIndex);
                    ChartTree.SelectedNode = chartContext.TreeNode;
                }
            }
        }

        internal void Initialise(
            ApiClientManager clientManager,
            ConsoleHandler1 consoleHandler,
            IHistoricalDataStore histDataStore,
            IContractStore contractStore) 
        {
            mClientManager = clientManager;
            mConsoleHandler = consoleHandler;
            mHistDataStore = histDataStore;

            var f = mClientManager.MarketDataFactory;
            mMarketDataManager = MD.CreateRealtimeDataManager(f,contractStore, null, mStudyLibraryManager);

        }

        internal void AddTreeEntry(
                        _IFuture contractFuture,
                        List<string> pathElements,
                        List<TimePeriod> timeframes,
                        TimePeriod initialTimeframe,
                        ChartSpecifier chartSpec,
                        bool showChart)
        {
            G.Assert(timeframes.Count != 0, "no timeframes");
            G.Assert(contractFuture != null, "no contract future");
            G.Assert(chartSpec != null, "no ChartSpec");

            G.Logger.Log($"Adding tree entry: from {chartSpec.FromTime}; to {chartSpec.ToTime}", nameof(AddTreeEntry), nameof(FChart));
            var first = (ChartTree.Nodes.Count == 0);
            var context = new ChartContext(
                                    createChartTreeNodes(ChartTree.Nodes, pathElements),
                                    contractFuture,
                                    timeframes,
                                    initialTimeframe,
                                    chartSpec,
                                    (s) => this.Text = s);
            if (first) ChartTree.SelectedNode = null;
            if (showChart)
            {
                AddChart(context);
            }




            TreeNodeCollection createChartTreeNodes(
                    TreeNodeCollection nodes,
                    List<string> pathEls)
            {
                TreeNodeCollection currNodes = nodes;
                foreach (string pathElenent in pathEls)
                {
                    if (!currNodes.ContainsKey(pathElenent))
                    {
                        var pathNode = currNodes.Add(pathElenent, pathElenent);
                        currNodes = pathNode.Nodes;
                    }
                    else
                    {
                        currNodes = currNodes[pathElenent].Nodes;
                    }
                }

                return currNodes;
            }

        }

        private void AddChart(
                        ChartContext context)
        {
            G.Logger.Log("Creating market data source", nameof(AddChart), nameof(ShowChartProcessor));
            context.DataSource = mMarketDataManager.CreateMarketDataSource(context.ContractFuture, false, "", false, null);
            context.DataSource.StartMarketData();

            context.MultiChart = createMultiChart();
            ChartTabPage tabPage = createTabPage(context);

            context.Ready += (s, e) =>
             {
                 tabPage.Controls.Add(context.MultiChart);

                 if (tabControlExtra1.TabPages.Count == 0)
                 {
                     mCurrentTabPage = tabPage;
                 }
                 tabControlExtra1.TabPages.Add(tabPage);

                 context.MultiChart.Visible = false;
                 context.MultiChart.Initialise(
                     createTimeframes(context.DataSource),
                     mHistDataStore.TimePeriodValidator,
                     context.ChartSpec,
                     ChartSkil.ChartStylesManager.Item(ChartStyles.ChartStyleNameBlack),
                     null,
                     "",
                     "",
                     false,
                     ColorTranslator.ToOle(Color.Black));

                 int initialIndex = 1;
                 foreach (TimePeriod t in context.Timeframes)
                 {
                     G.Logger.Log($"Adding timeframe {t.ToShortString()}", nameof(AddChart), nameof(FChart));
                     var index = context.MultiChart.Add(t, "", true, -1, false, true);
                     if (t == context.InitialTimeframe) initialIndex = index;
                 }
                 context.InitialChartIndex = initialIndex;
                 context.MultiChart.SelectChart(initialIndex);
                 tabControlExtra1.SelectedTab = tabPage;

                 this.ResumeLayout(true);
                 context.MultiChart.Visible = true;
             };

            context.Error += (s, e) =>
            {
                context.TreeNode.Parent?.Nodes.Remove(context.TreeNode);
                if (tabControlExtra1.TabPages.Contains(context.TabPage)) tabControlExtra1.TabPages.Remove(context.TabPage);
            };

            AxMultiChart createMultiChart()
            {
                G.Logger.Log("Creating multichart", nameof(createMultiChart), nameof(FChart));
                AxMultiChart mc = new AxMultiChart();
                ((System.ComponentModel.ISupportInitialize)(mc)).BeginInit();
                mc.Dock = DockStyle.Fill;
                mc.Enabled = true;
                mc.Location = new System.Drawing.Point(0, 0);
                mc.Name = "mc";
                mc.Size = new System.Drawing.Size(842, 516);
                mc.TabIndex = 0;
                ((System.ComponentModel.ISupportInitialize)(mc)).EndInit();
                return mc;
            }

            ChartTabPage createTabPage(ChartContext contxt)
            {
                G.Logger.Log("Creating tabpage", nameof(createTabPage), nameof(FChart));
                ChartTabPage t = new ChartTabPage(contxt);
                t.SuspendLayout();

                t.Location = new System.Drawing.Point(4, 30);
                t.Padding = new Padding(0);
                t.Size = new System.Drawing.Size(this.Width, this.Height);
                t.UseVisualStyleBackColor = true;

                t.ResumeLayout();
                context.TabPage = t;
                return t;
            }

            Timeframes createTimeframes(_IMarketDataSource dataSource)
            {
                return TF.CreateTimeframes(
                        dataSource.StudyBase,
                        dataSource.ContractFuture,
                        mHistDataStore,
                        Contracts.CreateClockFuture(dataSource.ContractFuture));
            }

        }

        internal void
        Sort()
        {
            ChartTree.Sort();
        }

        protected override void OnClosing(CancelEventArgs e)
        {
            TW.LogMessage("Main form closing");
            foreach (ChartTabPage t in tabControlExtra1.TabPages)
            {
                t.Context.DataSource?.Finish();
            }
            base.OnClosing(e);
        }

        static string
        GetContractName(_IContractSpecifier contractSpec)
        {
            return $"{contractSpec.LocalSymbol}@{contractSpec.Exchange}";
        }

        private class ChartContext : IGenericTickListener, IStateChangeListener
        {
            private _IMarketDataSource mDataSource;
            internal _IMarketDataSource
            DataSource
            {
                get => mDataSource;
                set
                {
                    mDataSource = value;
                    mDataSource.AddGenericTickListener(this);
                    mDataSource.AddStateChangeListener(this);

                    getInitialTickerValues();
                }
            }

            private readonly FutureWaiter
            mFutureWaiter = new FutureWaiter();

            internal _IFuture
            ContractFuture
            { get; private set; }

            internal IContract
            Contract { get; private set; }

            internal AxMultiChart
            MultiChart
            { get; set; }

            internal int
            InitialChartIndex { get; set; }

            internal bool
            IsCurrent { private get; set; }

            internal string
            Name { get; private set; }

            internal ChartTabPage
            TabPage { get; set; }

            internal ChartTreeNode
            TreeNode { get; private set; }

            internal List<TimePeriod>
            Timeframes { get; private set; }

            internal TimePeriod
            InitialTimeframe { get; private set; }

            internal ChartSpecifier
            ChartSpec { get; private set; }

            private TreeNodeCollection
            mParentTreeNodes;

            private readonly Action<string>
            mSetCaption;

            private SecurityTypes mSecType;
            private double mTickSize;

            private string mCurrentBid;
            private string mCurrentAsk;
            private string mCurrentTrade;
            private string mCurrentVolume;
            private string mCurrentHigh;
            private string mCurrentLow;
            private string mPreviousClose;

            internal ChartContext(
                    TreeNodeCollection parentTreeNodes,
                    _IFuture contractFuture,
                    List<TimePeriod> timeframes,
                    TimePeriod initialTimeframe,
                    ChartSpecifier chartSpec,
                    Action<string> setCaption)
            {
                mParentTreeNodes = parentTreeNodes;
                ContractFuture = contractFuture;
                Timeframes = timeframes;
                InitialTimeframe = initialTimeframe;
                ChartSpec = chartSpec;
                mSetCaption = setCaption;

                mFutureWaiter.WaitCompleted += mFutureWaiter_WaitCompleted;
                mFutureWaiter.Add(contractFuture);

            }

            internal event EventHandler<EventArgs>
            Error;
            protected virtual void OnError(EventArgs e)
            {
                TW.LogMessage("Chart context error");
                Error(this, e);
            }

            internal event EventHandler<EventArgs>
            Ready;
            protected virtual void OnReady(EventArgs e)
            {
                TW.LogMessage("Chart context ready");
                Ready(this, e);
            }

            private void mFutureWaiter_WaitCompleted(ref FutureWaitCompletedEventData ev)
            {
                if (!ev.Future.IsAvailable) return;
                Contract = ev.Future.Value as IContract;
                mSecType = Contract.Specifier.SecType;
                mTickSize = Contract.TickSize;

                var TreeNode = new ChartTreeNode(this, FChart.GetContractName(Contract.Specifier));
                mParentTreeNodes.Add(TreeNode);
                TreeNode.EnsureVisible();
            }


            void _IGenericTickListener.NoMoreTicks(ref GenericTickEventData ev) { }

            void _IGenericTickListener.NotifyTick(ref GenericTickEventData ev)
            {
                try
                {
                    switch (ev.Tick.TickType)
                    {
                        case TickTypes.TickTypeBid:
                            mCurrentBid = getFormattedPrice(ev.Tick.Price);
                            break;
                        case TickTypes.TickTypeAsk:
                            mCurrentAsk = getFormattedPrice(ev.Tick.Price);
                            break;
                        case TickTypes.TickTypeClosePrice:
                            mPreviousClose = getFormattedPrice(ev.Tick.Price);
                            break;
                        case TickTypes.TickTypeHighPrice:
                            mCurrentHigh = getFormattedPrice(ev.Tick.Price);
                            break;
                        case TickTypes.TickTypeLowPrice:
                            mCurrentLow = getFormattedPrice(ev.Tick.Price);
                            break;
                        case TickTypes.TickTypeTrade:
                            mCurrentTrade = getFormattedPrice(ev.Tick.Price);
                            break;
                        case TickTypes.TickTypeVolume:
                            mCurrentVolume = ev.Tick.Size.ToString();
                            break;
                        default:
                            return;
                    }

                    if (IsCurrent) mSetCaption(GetCaption(MultiChart));
                }
                catch (Exception ex)
                {
                    string inner = ex.InnerException?.StackTrace + ex.InnerException != null ? Environment.NewLine : String.Empty;
                    throw new System.Runtime.InteropServices.COMException(ex.Message)
                    {
                        Source = $"{inner}{ex.Source}{Environment.NewLine}{ex.StackTrace}"
                    };
                }
            }

            void _IStateChangeListener.Change(ref StateChangeEventData ev)
            {
                try
                {
                    var dataSource = ev.Source as IMarketDataSource;
                    if (ev.State == (int)MarketDataSourceStates.MarketDataSourceStateReady)
                    {
                        var contract = (IContract)dataSource.ContractFuture.Value;
                        Name = FChart.GetContractName(contract.Specifier);
                        dataSource.AddGenericTickListener(this);
                        OnReady(EventArgs.Empty);
                    }
                    else if (ev.State == (int)MarketDataSourceStates.MarketDataSourceStateRunning)
                    {
                        getInitialTickerValues();
                    }
                    else if (ev.State == (int)MarketDataSourceStates.MarketDataSourceStateStopped ||
                                ev.State == (int)MarketDataSourceStates.MarketDataSourceStateFinished)
                    {
                        // the ticker has been stopped before the chart has been closed
                        MultiChart.Finish();
                    }
                    else if (ev.State == (int)MarketDataSourceStates.MarketDataSourceStateError)
                    {
                        OnError(EventArgs.Empty);
                    }
                }
                catch (Exception ex)
                {
                    string inner = ex.InnerException?.StackTrace + ex.InnerException != null ? Environment.NewLine : String.Empty;
                    throw new System.Runtime.InteropServices.COMException(ex.Message)
                    {
                        Source = $"{inner}{ex.Source}{Environment.NewLine}{ex.StackTrace}"
                    };
                }
            }

            internal string
            GetCaption(AxMultiChart multiChart)
            {
                string s;
                if (multiChart.Count == 0)
                {
                    s = Contract.Specifier.LocalSymbol;
                }
                else
                {
                    s = $"{Contract.Specifier.LocalSymbol} ({multiChart.get_TimePeriod().ToString()})";
                }
                return $"{s}    B={mCurrentBid}  T={mCurrentTrade}  A={mCurrentAsk}  V={mCurrentVolume}  H={mCurrentHigh}  L={mCurrentLow}  C={mPreviousClose}";
            }

            private string getFormattedPrice(double pPrice)
            {
                return Contracts.FormatPrice(pPrice, mSecType, mTickSize);
            }

            private void
            getInitialTickerValues()
            {
                if (DataSource.State != MarketDataSourceStates.MarketDataSourceStateRunning) return;
                mCurrentBid = getFormattedPrice(DataSource.CurrentQuote[TickTypes.TickTypeBid].Price);
                mCurrentTrade = getFormattedPrice(DataSource.CurrentQuote[TickTypes.TickTypeTrade].Price);
                mCurrentAsk = getFormattedPrice(DataSource.CurrentQuote[TickTypes.TickTypeAsk].Price);
                mCurrentVolume = DataSource.CurrentQuote[TickTypes.TickTypeVolume].Size.ToString();
                mCurrentHigh = getFormattedPrice(DataSource.CurrentQuote[TickTypes.TickTypeHighPrice].Price);
                mCurrentLow = getFormattedPrice(DataSource.CurrentQuote[TickTypes.TickTypeLowPrice].Price);
                mPreviousClose = getFormattedPrice(DataSource.CurrentQuote[TickTypes.TickTypeClosePrice].Price);
            }

        }

        private class ChartTabPage : TabPage
        {
            internal ChartContext
            Context { get; private set; }

            internal ChartTabPage(ChartContext context): base()
            {
                Context = context;
                Context.Ready += (s, e) =>
                {
                    Name = FChart.GetContractName(context.Contract.Specifier);
                    BackColor = Color.Black;
                    Text = Name;
                };
            }

        }

        private class ChartTreeNode : TreeNode
        {
            internal ChartContext
            Context
            { get; private set; }

            internal ChartTreeNode(
                ChartContext context,
                string text) : base(text)
            {
                Context = context;
                Name = Text;
            }

        }
    }
}
