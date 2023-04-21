using Autodesk.Revit.DB;
using Prism.Mvvm;
using RebarsToExcel.Commands;
using RebarsToExcel.Models;
using RebarsToExcel.Models.Bars;
using RebarsToExcel.Views;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Windows.Data;

namespace RebarsToExcel.ViewModels
{
    public class DataManageVM : BindableBase
    {
        private readonly string _beamGroupModelParamValue = "Балки монолитные";
        private readonly string _columnGroupModelParamValue = "Колонны монолитные";
        private readonly string _floorGroupModelParamValue = "Перекрытия монолитные";
        private readonly string _wallGroupModelParamValue = "Стены монолитные";
        private readonly string _rebarGroupModelParamValue = "Детали";
        private readonly string _assemblyGroupModelParamValue = "Сборочные единицы";
        private readonly string _selectAll = "(все)";

        private readonly string _rebarClassParamName = "_Класс арматуры";
        private readonly string _rebarDiameterParamName = "_Диаметр стержня";
        private readonly string _rebarMassParamName = "_Масса";
        private readonly string _rebarMassPerLengthParamName = "_Масса на ед. длины";
        private readonly string _rebarShapeParamName = "_Форма стержня";
        private readonly string _rebarShapeImageParamName = "_Изображение формы";
        private readonly string _rebarLengthParamName = "_Длина стержня";
        private readonly string _rebarLengthСoefficientParamName = "_Коэф. перепуска";
        private readonly string _rebarCountTypeParamName = "_Тип подсчёта количества";
        private readonly string _rebarCountParamName = "_Количество";
        private readonly string _rebarMarkParamName = "_Марка";
        private readonly string _rebarDefinitionParamName = "_Обозначение";
        private readonly string _rebarNominationParamName = "_Наименование";
        private readonly string _rebarTypeOfConstructionParamName = "_Тип основы";
        private readonly string _rebarMarkOfConstructionParamName = "_Метка основы";
        private readonly string _rebarConstructionCountParamName = "_Количество основ";
        private readonly string _rebarTypicalFloorParamName = "_Типовой этаж";
        private readonly string _rebarTypicalFloorCountParamName = "_Количество типовых этажей";
        private readonly string _rebarLevelParamName = "_Этаж";
        private readonly string _rebarSectionParamName = "_Секция";
        private readonly string _projectNameParamName = "_Наименование объекта";

        private Guid _rebarLengthParamGuid = Guid.Empty;
        private Guid _rebarMarkParamGuid = Guid.Empty;
        private Guid _rebarDefinitionParamGuid = Guid.Empty;
        private Guid _rebarNominationParamGuid = Guid.Empty;
        private Guid _rebarTypeOfConstructionParamGuid = Guid.Empty;
        private Guid _rebarMarkOfConstructionParamGuid = Guid.Empty;
        private Guid _rebarConstructionCountParamGuid = Guid.Empty;
        private Guid _rebarTypicalFloorParamGuid = Guid.Empty;
        private Guid _rebarTypicalFloorCountParamGuid = Guid.Empty;
        private Guid _rebarLevelParamGuid = Guid.Empty;
        private Guid _rebarSectionParamGuid = Guid.Empty;

        private readonly IDictionary<BarSize, double> _sizes = new Dictionary<BarSize, double>
        {
            { new BarSize(0, "_A"), 0 },
            { new BarSize(0, "_Aокр"), 0 },
            { new BarSize(1, "_B"), 0 },
            { new BarSize(1, "_Bокр"), 0 },
            { new BarSize(2, "_C"), 0 },
            { new BarSize(2, "_Cокр"), 0 },
            { new BarSize(3, "_D"), 0 },
            { new BarSize(3, "_Dокр"), 0 },
            { new BarSize(4, "_E"), 0 },
            { new BarSize(4, "_Eокр"), 0 },
            { new BarSize(5, "_F"), 0 },
            { new BarSize(6, "_G"), 0 },
            { new BarSize(7, "_H"), 0 },
            { new BarSize(8, "_J"), 0 },
            { new BarSize(9, "_K"), 0 },
            { new BarSize(10, "_Угол α"), 0 },
            { new BarSize(11, "_Угол β"), 0 },
            { new BarSize(12, "_Угол γ"), 0 },
            { new BarSize(13, "_Угол δ"), 0 },
        };

        private Document _doc;
        private IList<Element> _levels;

        public bool IsDataEmpty { get; private set; }
        public string DocumentTitle { get; private set; }
        public double BarsTotalCount { get; private set; }
        public double RebarAssembliesTotalCount { get; private set; }
        public BackgroundWorker BackgroundWorker = new BackgroundWorker();

        #region СВОЙСТВА ДЕТАЛЕЙ
        private RebarLevel _selectedBarLevel;
        public RebarLevel SelectedBarLevel
        {
            get => _selectedBarLevel;
            set
            {
                _selectedBarLevel = value;
                RaisePropertyChanged(nameof(SelectedBarLevel));
            }
        }

        private string _selectedBarSection;
        public string SelectedBarSection
        {
            get => _selectedBarSection;
            set
            {
                _selectedBarSection = value;
                FilterBarElements();
                FilterBarConstructionMarks();
                RaisePropertyChanged(nameof(SelectedBarSection));
            }
        }

        private string _selectedBarConstructionType;
        public string SelectedBarConstructionType
        {
            get => _selectedBarConstructionType;
            set
            {
                _selectedBarConstructionType = value;
                FilterBarElements();
                FilterBarConstructionMarks();
                RaisePropertyChanged(nameof(SelectedBarConstructionType));
            }
        }

        private string _selectedBarConstructionMark;
        public string SelectedBarConstructionMark
        {
            get => _selectedBarConstructionMark;
            set
            {
                _selectedBarConstructionMark = value;
                FilterBarElements();
                RaisePropertyChanged(nameof(SelectedBarConstructionMark));
            }
        }

        public IList<string> BarSections { get; set; }
        public IList<string> BarConstructionTypes { get; set; }

        private IList<string> _barConstructionMarks;
        public IList<string> BarConstructionMarks
        {
            get => _barConstructionMarks;
            set
            {
                _barConstructionMarks = value;
                RaisePropertyChanged(nameof(BarConstructionMarks));
            }
        }

        public IList<RebarLevel> BarLevels { get; set; }
        public ICollectionView BarsCollectionView { get; set; }

        private Bar _selectedBar;
        public Bar SelectedBar 
        {
            get => _selectedBar;
            set
            {
                _selectedBar = value;
                RaisePropertyChanged(nameof(SelectedBar));
            }
        }

        public string SelectedBarShapeImagePath
        {
            get
            {
                if (_selectedBar == null)
                    return null;

                return _selectedBar.ShapeImagePath;
            }
        }
        #endregion

        #region СВОЙСТВА СБОРОЧНЫХ ЕДИНИЦ
        private RebarLevel _selectedRebarAssemblyLevel;
        public RebarLevel SelectedRebarAssemblyLevel
        {
            get => _selectedRebarAssemblyLevel;
            set
            {
                _selectedRebarAssemblyLevel = value;
                RaisePropertyChanged(nameof(SelectedRebarAssemblyLevel));
            }
        }

        private string _selectedRebarAssemblySection;
        public string SelectedRebarAssemblySection
        {
            get => _selectedRebarAssemblySection;
            set
            {
                _selectedRebarAssemblySection = value;
                FilterRebarAssemblyElements();
                FilterRebarAssemblyConstructionMarks();
                RaisePropertyChanged(nameof(SelectedRebarAssemblySection));
            }
        }

        private string _selectedRebarAssemblyConstructionType;
        public string SelectedRebarAssemblyConstructionType
        {
            get => _selectedRebarAssemblyConstructionType;
            set
            {
                _selectedRebarAssemblyConstructionType = value;
                FilterRebarAssemblyElements();
                FilterRebarAssemblyConstructionMarks();
                RaisePropertyChanged(nameof(SelectedRebarAssemblyConstructionType));
            }
        }

        private string _selectedRebarAssemblyConstructionMark;
        public string SelectedRebarAssemblyConstructionMark
        {
            get => _selectedRebarAssemblyConstructionMark;
            set
            {
                _selectedRebarAssemblyConstructionMark = value;
                FilterRebarAssemblyElements();
                RaisePropertyChanged(nameof(SelectedRebarAssemblyConstructionMark));
            }
        }
        
        public IList<string> RebarAssemblySections { get; set; }
        public IList<string> RebarAssemblyConstructionTypes { get; set; }

        private IList<string> _rebarAssemblyConstructionMarks;
        public IList<string> RebarAssemblyConstructionMarks 
        { 
            get => _rebarAssemblyConstructionMarks; 
            set
            {
                _rebarAssemblyConstructionMarks = value;
                RaisePropertyChanged(nameof(RebarAssemblyConstructionMarks));
            }
        }

        public IList<RebarLevel> RebarAssemblyLevels { get; set; }
        public ICollectionView RebarAssembliesCollectionView { get; set; }
        public RebarAssembly SelectedRebarAssembly { get; set; }
        #endregion

        #region СВОЙСТВА PROGRESSBAR
        private double _barsProgressCounter;
        public double BarsProgressCounter
        {
            get => _barsProgressCounter;
            set
            {
                _barsProgressCounter = value;
                RaisePropertyChanged(nameof(BarsProgressCounter));
            }
        }

        private double _assembliesProgressCounter;
        public double RebarAssembliesProgressCounter
        {
            get => _assembliesProgressCounter;
            set
            {
                _assembliesProgressCounter = value;
                RaisePropertyChanged(nameof(RebarAssembliesProgressCounter));
            }
        }
        #endregion

        #region COMMANDS
        private RelayCommand _selectBarLevelCommand;
        public RelayCommand SelectBarLevelCommand
        {
            get
            {
                return _selectBarLevelCommand ?? (_selectBarLevelCommand = new RelayCommand(obj =>
                {
                    FilterBarElements();
                    FilterBarConstructionMarks();
                }));
            }
        }

        private RelayCommand _selectAllBarLevelsCommand;
        public RelayCommand SelectAllBarLevelsCommand
        {
            get
            {
                return _selectAllBarLevelsCommand ?? (_selectAllBarLevelsCommand = new RelayCommand(obj =>
                {
                    SelectAllBarLevels();
                    FilterBarElements();
                    FilterBarConstructionMarks();
                }));
            }
        }

        private RelayCommand _unselectAllBarLevelsCommand;
        public RelayCommand UnselectAllBarLevelsCommand
        {
            get
            {
                return _unselectAllBarLevelsCommand ?? (_unselectAllBarLevelsCommand = new RelayCommand(obj =>
                {
                    UnselectAllBarLevels();
                    FilterBarElements();
                    FilterBarConstructionMarks();
                }));
            }
        }
        private RelayCommand _selectRebarAssemblyLevelCommand;
        public RelayCommand SelectRebarAssemblyLevelCommand
        {
            get
            {
                return _selectRebarAssemblyLevelCommand ?? (_selectRebarAssemblyLevelCommand = new RelayCommand(obj => 
                { 
                    FilterRebarAssemblyElements();
                    FilterRebarAssemblyConstructionMarks();
                }));
            }
        }

        private RelayCommand _selectAllRebarAssemblyLevelsCommand;
        public RelayCommand SelectAllRebarAssemblyLevelsCommand
        {
            get
            {
                return _selectAllRebarAssemblyLevelsCommand ?? (_selectAllRebarAssemblyLevelsCommand = new RelayCommand(obj =>
                {
                    SelectAllRebarAssemblyLevels();
                    FilterRebarAssemblyElements();
                    FilterRebarAssemblyConstructionMarks();
                }));
            }
        }

        private RelayCommand _unselectAllRebarAssemblyLevelsCommand;
        public RelayCommand UnselectAllRebarAssemblyLevelsCommand
        {
            get
            {
                return _unselectAllRebarAssemblyLevelsCommand ?? (_unselectAllRebarAssemblyLevelsCommand = new RelayCommand(obj => 
                {
                    UnselectAllRebarAssemblyLevels();
                    FilterRebarAssemblyElements();
                    FilterRebarAssemblyConstructionMarks();
                }));
            }
        }

        private RelayCommand _exportToExcelCommand;
        public RelayCommand ExportToExcelCommand
        {
            get
            {
                return _exportToExcelCommand ?? (_exportToExcelCommand = new RelayCommand(obj => ExportToExcel()));
            }
        }
        #endregion

        public DataManageVM(Document doc)
        {
            _doc = doc;
            DocumentTitle = _doc.Title;
            RebarsToExcelApp.IsOpened = true;

            SetData();
        }

        #region ЗАПИСЬ ВСЕХ ДАННЫХ
        private void SetData()
        {
            _levels = GetSystemLevels();

            var typicalFloors = GetAllTypicalFloors();
            var allBarElements = GetAllBars();
            var allGenericModelElements = GetAllGenericModelElements();
            var allSystemAssemblyElements = GetAllSystemAssemblyElements();

            BarsTotalCount = allBarElements.Count;
            BarsProgressCounter = 0;
            RebarAssembliesTotalCount = allGenericModelElements.Count + allSystemAssemblyElements.Count;
            RebarAssembliesProgressCounter = 0;

            if (BarsTotalCount == 0 && RebarAssembliesTotalCount == 0)
            {
                IsDataEmpty = true;
                return;
            }

            AnalyzeWindow analyzeWindow = new AnalyzeWindow(this);

            BackgroundWorker.DoWork += (sender, e) =>
            {
                AnalyzeAllBars(allBarElements, typicalFloors);
                AnalyzeAllRebarAssemblies(allGenericModelElements, allSystemAssemblyElements, typicalFloors);
            };

            BackgroundWorker.RunWorkerCompleted += (sender, e) =>
            {
                if (analyzeWindow != null)
                {
                    analyzeWindow.Close();
                }

                //Получение деталей
                var barsData = BarsData.GetData();
                BarsCollectionView = CollectionViewSource.GetDefaultView(barsData);

                BarLevels = BarsData.GetLevels();
                BarSections = BarsData.GetSections();
                BarSections.Insert(0, _selectAll);
                BarConstructionTypes = BarsData.GetConstructionTypes();
                BarConstructionTypes.Insert(0, _selectAll);

                SelectedBarSection = BarSections.FirstOrDefault();
                SelectedBarConstructionType = BarConstructionTypes.FirstOrDefault();
                SelectedBarConstructionMark = _selectAll;

                //Получение сборочных единиц
                var rebarAssembliesData = RebarAssembliesData.GetData();
                RebarAssembliesCollectionView = CollectionViewSource.GetDefaultView(rebarAssembliesData);

                RebarAssemblyLevels = RebarAssembliesData.GetLevels();
                RebarAssemblySections = RebarAssembliesData.GetSections();
                RebarAssemblySections.Insert(0, _selectAll);
                RebarAssemblyConstructionTypes = RebarAssembliesData.GetConstructionTypes();
                RebarAssemblyConstructionTypes.Insert(0, _selectAll);

                SelectedRebarAssemblySection = RebarAssemblySections.FirstOrDefault();
                SelectedRebarAssemblyConstructionType = RebarAssemblyConstructionTypes.FirstOrDefault();
                SelectedRebarAssemblyConstructionMark = _selectAll;

                FilterRebarAssemblyElements();
            };

            analyzeWindow.Show();

            BackgroundWorker.RunWorkerAsync();
        }
        #endregion

        #region ВЫБРАТЬ И ОТМЕНИТЬ ВЫБОР
        private void SelectAllBarLevels()
        {
            foreach (var level in BarLevels)
            {
                level.IsSelected = true;
            }
        }

        private void UnselectAllBarLevels()
        {
            foreach (var level in BarLevels)
            {
                level.IsSelected = false;
            }
        }

        private void SelectAllRebarAssemblyLevels()
        {
            foreach (var level in RebarAssemblyLevels)
            {
                level.IsSelected = true;
            }
        }

        private void UnselectAllRebarAssemblyLevels()
        {
            foreach (var level in RebarAssemblyLevels)
            {
                level.IsSelected = false;
            }
        }
        #endregion

        #region ФИЛЬТРАЦИЯ ТАБЛИЦ
        private void FilterBarElements()
        {
            var selectedLevels = BarLevels.Where(level => level.IsSelected).Select(level => level.Name).ToList();
            var selectionAll = "(все)";

            BarsCollectionView.Filter = bar => selectedLevels.Contains((bar as Bar).Level.Name)
            && SelectedBarSection.Contains(SelectedBarSection == selectionAll ? string.Empty : (bar as Bar).Section)
            && SelectedBarConstructionType.Contains(SelectedBarConstructionType == selectionAll ? string.Empty : (bar as Bar).ConstructionType)
            && (SelectedBarConstructionMark == selectionAll ? SelectedBarConstructionMark.Contains(string.Empty) : SelectedBarConstructionMark == (bar as Bar).ConstructionMark);
        }

        private void FilterBarConstructionMarks()
        {
            var bars = BarsCollectionView.Cast<Bar>().ToList();
            var constructionMarks = bars.Select(bar => bar.ConstructionMark).Distinct().OrderBy(mark => mark).ToList();
            constructionMarks.Insert(0, "(все)");
            BarConstructionMarks = constructionMarks;

            SelectedBarConstructionMark = BarConstructionMarks.FirstOrDefault(mark => mark == SelectedBarConstructionMark) ?? "(все)";
        }

        private void FilterRebarAssemblyElements()
        {
            var selectedLevels = RebarAssemblyLevels.Where(level => level.IsSelected).Select(level => level.Name).ToList();
            var selectionAll = "(все)";

            RebarAssembliesCollectionView.Filter = rebarAssembly => selectedLevels.Contains((rebarAssembly as RebarAssembly).Level.Name)
            && SelectedRebarAssemblySection.Contains(SelectedRebarAssemblySection == selectionAll ? string.Empty : (rebarAssembly as RebarAssembly).Section)
            && SelectedRebarAssemblyConstructionType.Contains(SelectedRebarAssemblyConstructionType == selectionAll ? string.Empty : (rebarAssembly as RebarAssembly).ConstructionType)
            && (SelectedRebarAssemblyConstructionMark == selectionAll? SelectedRebarAssemblyConstructionMark.Contains(string.Empty) : SelectedRebarAssemblyConstructionMark == (rebarAssembly as RebarAssembly).ConstructionMark);
        }

        private void FilterRebarAssemblyConstructionMarks()
        {
            var rebarAssemblies = RebarAssembliesCollectionView.Cast<RebarAssembly>().ToList();
            var constructionMarks = rebarAssemblies.Select(rebarAssembly => rebarAssembly.ConstructionMark).Distinct().OrderBy(mark => mark).ToList();
            constructionMarks.Insert(0, "(все)");
            RebarAssemblyConstructionMarks = constructionMarks;

            SelectedRebarAssemblyConstructionMark = RebarAssemblyConstructionMarks.FirstOrDefault(mark => mark == SelectedRebarAssemblyConstructionMark) ?? "(все)";
        }
        #endregion

        #region ПОЛУЧЕНИЕ КОНСТРУКЦИЙ
        private List<Element> GetAllConstructions()
        {
            var categories = new List<BuiltInCategory> 
            { 
                BuiltInCategory.OST_StructuralFraming,
                BuiltInCategory.OST_StructuralColumns,
                BuiltInCategory.OST_Floors,
                BuiltInCategory.OST_Walls 
            };

            var categoriesCollection = new List<BuiltInCategory>(categories);
            var categoriesFilter = new ElementMulticategoryFilter(categoriesCollection);

            return new FilteredElementCollector(_doc)
                .WherePasses(categoriesFilter)
                .WhereElementIsNotElementType()
                .ToElements()
                .ToList();
        }

        private List<Element> GetAllBeams(List<Element> constructions)
        {
            return constructions.Where(x => GetElementType(x).get_Parameter(BuiltInParameter.ALL_MODEL_MODEL)?.AsString() == _beamGroupModelParamValue).ToList();
        }

        private List<Element> GetAllColumns(List<Element> constructions)
        {
            return constructions.Where(x => GetElementType(x).get_Parameter(BuiltInParameter.ALL_MODEL_MODEL)?.AsString() == _columnGroupModelParamValue).ToList();
        }

        private List<Element> GetAllFloors(List<Element> constructions)
        {
            return constructions.Where(x => GetElementType(x).get_Parameter(BuiltInParameter.ALL_MODEL_MODEL)?.AsString() == _floorGroupModelParamValue).ToList();
        }

        private List<Element> GetAllWalls(List<Element> constructions)
        {
            return constructions.Where(x => GetElementType(x).get_Parameter(BuiltInParameter.ALL_MODEL_MODEL)?.AsString() == _wallGroupModelParamValue).ToList();
        }
        #endregion

        #region ПОЛУЧЕНИЕ ТИПОВЫХ ЭТАЖЕЙ
        private IList<TypicalFloor> GetAllTypicalFloors()
        {
            var constructions = GetAllConstructions();
            var beams = GetAllBeams(constructions);
            var columns = GetAllColumns(constructions);
            var floors = GetAllFloors(constructions);
            var walls = GetAllWalls(constructions);

            var beamTypicalFloors = GetAllBeamTypicalFloors(beams);
            var columnTypicalFloors = GetAllColumnTypicalFloors(columns);
            var floorTypicalFloors = GetAllFloorTypicalFloors(floors);
            var wallTypicalFloors = GetAllWallTypicalFloors(walls);

            var typicalFloors = beamTypicalFloors.Concat(columnTypicalFloors)
                .Concat(floorTypicalFloors)
                .Concat(wallTypicalFloors)
                .ToList();

            return typicalFloors;
        }

        private IList<TypicalFloor> GetAllBeamTypicalFloors(IList<Element> beams)
        {
            var beamsWithMultipleTypicalFloors = beams.Where(element => GetTypicalFloorCount(element) > 1).ToList();

            var beamTypicalLevels = beamsWithMultipleTypicalFloors.GroupBy(element => GetTypicalFloor(element))
                .Select(element => new TypicalFloor(ConstructionType.Beam, element.Key, element.Select(el => GetLevel(el)).Distinct(new LevelComparer()).ToList()))
                .ToList();

            return beamTypicalLevels;
        }

        private IList<TypicalFloor> GetAllColumnTypicalFloors(IList<Element> columns)
        {
            var columnsWithMultipleTypicalFloors = columns.Where(element => GetTypicalFloorCount(element) > 1).ToList();

            var columnTypicalLevels = columnsWithMultipleTypicalFloors.GroupBy(element => GetTypicalFloor(element))
                .Select(element => new TypicalFloor(ConstructionType.Column, element.Key, element.Select(el => GetLevel(el)).Distinct(new LevelComparer()).ToList()))
                .ToList();

            return columnTypicalLevels;
        }

        private IList<TypicalFloor> GetAllFloorTypicalFloors(IList<Element> floors)
        {
            var floorsWithMultipleTypicalFloors = floors.Where(element => GetTypicalFloorCount(element) > 1).ToList();

            var floorTypicalLevels = floorsWithMultipleTypicalFloors.GroupBy(element => GetTypicalFloor(element))
                .Select(element => new TypicalFloor(ConstructionType.Floor, element.Key, element.Select(el => GetLevel(el)).Distinct(new LevelComparer()).ToList()))
                .ToList();

            return floorTypicalLevels;
        }

        private IList<TypicalFloor> GetAllWallTypicalFloors(IList<Element> walls)
        {
            var wallsWithMultipleTypicalFloors = walls.Where(element => GetTypicalFloorCount(element) > 1).ToList();

            var wallTypicalLevels = wallsWithMultipleTypicalFloors.GroupBy(element => GetTypicalFloor(element))
                .Select(element => new TypicalFloor(ConstructionType.Wall, element.Key, element.Select(el => GetLevel(el)).Distinct(new LevelComparer()).ToList()))
                .ToList();

            return wallTypicalLevels;
        }
        #endregion

        #region ПОЛУЧЕНИЕ И АНАЛИЗ АРМАТУРЫ
        private IList<Element> GetAllBars()
        {
            ParameterValueProvider groupModelProvider = new ParameterValueProvider(new ElementId(BuiltInParameter.ALL_MODEL_MODEL));
            FilterStringRuleEvaluator groupModelEvaluator = new FilterStringEquals();
            FilterRule groupModelFilterRule = new FilterStringRule(groupModelProvider, groupModelEvaluator, _rebarGroupModelParamValue, false);
            ElementParameterFilter groupModelFilter = new ElementParameterFilter(groupModelFilterRule);

            var rebarsAsClass = new FilteredElementCollector(_doc)
                .OfCategory(BuiltInCategory.OST_Rebar)
                .OfClass(typeof(Autodesk.Revit.DB.Structure.Rebar))
                .WhereElementIsNotElementType()
                .WherePasses(groupModelFilter)
                .ToElements()
                .ToList();

            var rebarsAsFamilyInstance = new FilteredElementCollector(_doc)
                .OfCategory(BuiltInCategory.OST_Rebar)
                .OfClass(typeof(FamilyInstance))
                .WhereElementIsNotElementType()
                .WherePasses(groupModelFilter)
                .ToElements()
                .Where(element => (element as FamilyInstance).SuperComponent == null)
                .ToList();

            rebarsAsClass.AddRange(rebarsAsFamilyInstance);
            return rebarsAsClass;
        }

        private IList<Element> GetAllGenericModelElements()
        {
            ParameterValueProvider provider = new ParameterValueProvider(new ElementId(BuiltInParameter.ALL_MODEL_MODEL));
            FilterStringRuleEvaluator evaluator = new FilterStringEquals();
            FilterRule groupModelFilterRule = new FilterStringRule(provider, evaluator, _assemblyGroupModelParamValue, false);
            ElementParameterFilter groupModelFilter = new ElementParameterFilter(groupModelFilterRule);

            return new FilteredElementCollector(_doc)
                .OfCategory(BuiltInCategory.OST_GenericModel)
                .WhereElementIsNotElementType()
                .WherePasses(groupModelFilter)
                .ToElements()
                .ToList();
        }

        private IList<Element> GetAllSystemAssemblyElements()
        {
            ParameterValueProvider provider = new ParameterValueProvider(new ElementId(BuiltInParameter.ALL_MODEL_MODEL));
            FilterStringRuleEvaluator evaluator = new FilterStringEquals();
            FilterRule groupModelFilterRule = new FilterStringRule(provider, evaluator, _assemblyGroupModelParamValue, false);
            ElementParameterFilter groupModelFilter = new ElementParameterFilter(groupModelFilterRule);

            return new FilteredElementCollector(_doc)
                .OfCategory(BuiltInCategory.OST_Assemblies)
                .WhereElementIsNotElementType()
                .WherePasses(groupModelFilter)
                .ToElements()
                .ToList();
        }

        private void AnalyzeAllBars(IList<Element> allBarElements, IList<TypicalFloor> typicalFloors)
        {
            foreach (var barElement in allBarElements) //Перебираем всю арматуру
            {
                var elementType = GetElementType(barElement);

                var barClassParamValue = GetClass(barElement, elementType); //_Класс
                var barDiameterParamValue = GetDiameter(barElement, elementType); //_Диаметр
                var barConstructionTypeParamValue = GetConstructionType(barElement); //_Тип основы
                var barCountTypeParamValue = GetCountType(elementType); //_Тип подсчета количества
                var barLengthСoefficient = GetBarLengthСoefficient(elementType); //_Коэф. перепуска
                var barCountParamValue = GetCount(barElement, elementType, barCountTypeParamValue, barLengthСoefficient); //_Количество
                var barLengthParamValue = GetBarLength(barElement, barCountTypeParamValue); //_Длина
                var barMassParamValue = GetBarMass(barElement, elementType, barCountTypeParamValue); //_Масса
                var barShapeParamValue = barElement.LookupParameter(_rebarShapeParamName).AsString(); //_Номер формы
                var positionParamValue = GetPosition(barElement); //Марка

                var bar = new Bar(barClassParamValue, barDiameterParamValue, barMassParamValue, barShapeParamValue)
                {
                    Id = barElement.Id,
                    Position = positionParamValue, //Марка
                    PositionWithShapeMark = string.Concat(positionParamValue, barShapeParamValue == "0.1" || barShapeParamValue == "0.2" ? string.Empty : "*"),
                    Length = barLengthParamValue,
                    ShapeImagePath = GetShapeImagePath(barElement),
                    CountType = barCountTypeParamValue,
                    CountTypeInfo = GetCountTypeInfo(barCountTypeParamValue),
                    Count = barCountParamValue,
                    Level = GetLevel(barElement),
                    Section = GetSection(barElement), //_Секция
                    ConstructionType = barConstructionTypeParamValue,
                    ConstructionTypeEnum = GetConstructionTypeEnum(barConstructionTypeParamValue),
                    ConstructionMark = GetConstructionMark(barElement), //_Метка основы
                    ConstructionCount = GetConstructionCount(barElement), //_Количество основ
                    TypicalFloor = GetTypicalFloor(barElement), //_Типовой этаж
                    TypicalFloorCount = GetTypicalFloorCount(barElement), //_Количество типовых этажей
                    DiameterClassLengthInfo = barCountTypeParamValue == 2? $"⌀{barDiameterParamValue} {barClassParamValue}" : $"⌀{barDiameterParamValue} {barClassParamValue}, L={barLengthParamValue}",
                };

                if (bar.Shape != "0.1" && bar.Shape != "0.2")
                {
                    SetBarSizeParameters(ref bar, barElement);
                }

                BarsData.AddBar(bar);
                BarsProgressCounter++;
            }

            BarsData.AnalyzeDataByConstructionCount();
            BarsData.AnalyzeDataByTypicalFloorCount(typicalFloors);

            ShapeImagesData.SaveToFolder();
        }

        private void AnalyzeAllRebarAssemblies(IList<Element> allGenericModelElements, IList<Element> allSystemAssemblyElements, IList<TypicalFloor> typicalFloors)
        {
            foreach (var genericModelElement in allGenericModelElements) //Перебираем все Обобщенные модели
            {
                var elementType = GetElementType(genericModelElement);
                var descriptionTypeParamValue = GetDescription(elementType); //Описание
                var globalModelTypeParamValue = GetGroupModel(elementType); //Группа модели

                var markParamValue = GetMark(elementType); //Маркировка типоразмера, _Марка или _Наименование
                var massParamValue = GetMass(genericModelElement, elementType); //_Масса
                var constructionTypeParamValue = GetConstructionType(genericModelElement); //_Тип основы

                var rebarAssembly = new RebarAssembly(descriptionTypeParamValue, markParamValue, globalModelTypeParamValue, massParamValue)
                {
                    Id = genericModelElement.Id,
                    Definition = GetDefinition(elementType), //_Обозначение
                    ConstructionType = constructionTypeParamValue,
                    ConstructionTypeEnum = GetConstructionTypeEnum(constructionTypeParamValue),
                    ConstructionMark = GetConstructionMark(genericModelElement), //_Метка основы
                    ConstructionCount = GetConstructionCount(genericModelElement), //_Количество основ
                    TypicalFloor = GetTypicalFloor(genericModelElement), //_Типовой этаж
                    TypicalFloorCount = GetTypicalFloorCount(genericModelElement), //_Количество типовых этажей
                    Level = GetLevel(genericModelElement), //_Этаж
                    Section = GetSection(genericModelElement) //_Секция
                };

                var allRebarIdsOfAsssembly = (genericModelElement as FamilyInstance).GetSubComponentIds();
                AddAllRebarsOfAssemblyToRebarAssembly(rebarAssembly, allRebarIdsOfAsssembly); //Добавляем всю вложенную арматуру семейства в rebarAssembly

                RebarAssembliesData.AddRebarAssembly(rebarAssembly);
                RebarAssembliesProgressCounter++;
            }

            foreach (var assemblyElement in allSystemAssemblyElements) //Перебираем все Сборки
            {
                var elementType = GetElementType(assemblyElement);
                var descriptionTypeParamValue = GetDescription(elementType); //Описание
                var globalModelTypeParamValue = GetGroupModel(elementType); //Группа модели

                var markParamValue = GetMark(elementType); //Маркировка типоразмера или _Марка
                var massParamValue = GetMass(assemblyElement, elementType); //_Масса
                var constructionTypeParamValue = GetConstructionType(assemblyElement); //_Тип основы

                var rebarAssembly = new RebarAssembly(descriptionTypeParamValue, markParamValue, globalModelTypeParamValue, massParamValue)
                {
                    Id = assemblyElement.Id,
                    Definition = GetDefinition(elementType), //_Обозначение
                    ConstructionType = constructionTypeParamValue,
                    ConstructionTypeEnum = GetConstructionTypeEnum(constructionTypeParamValue),
                    ConstructionMark = GetConstructionMark(assemblyElement), //_Метка основы
                    ConstructionCount = GetConstructionCount(assemblyElement), //_Количество основ
                    TypicalFloor = GetTypicalFloor(assemblyElement), //_Типовой этаж
                    TypicalFloorCount = GetTypicalFloorCount(assemblyElement), //_Количество типовых этажей
                    Level = GetLevel(assemblyElement), //_Этаж
                    Section = GetSection(assemblyElement) //_Секция
                };

                var allRebarIdsOfAsssembly = (assemblyElement as AssemblyInstance).GetMemberIds();
                AddAllRebarsOfAssemblyToRebarAssembly(rebarAssembly, allRebarIdsOfAsssembly); //Добавляем всю вложенную арматуру семейства в rebarAssembly

                RebarAssembliesData.AddRebarAssembly(rebarAssembly);
                RebarAssembliesProgressCounter++;
            }

            //Анализ сборочных единиц по количеству основ и типовым этажам
            RebarAssembliesData.AnalyzeDataByConstructionCount();
            RebarAssembliesData.AnalyzeDataByTypicalFloorCount(typicalFloors);
        }
        #endregion

        #region ПОЛУЧЕНИЕ ПАРАМЕТРОВ
        private string GetDescription(Element elementType)
        {
            return elementType.get_Parameter(BuiltInParameter.ALL_MODEL_DESCRIPTION)?.AsString();
        }

        private string GetGroupModel(Element elementType)
        {
            return elementType.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL)?.AsString();
        }

        private string GetPosition(Element element)
        {
            var markParam = element.get_Parameter(BuiltInParameter.ALL_MODEL_MARK);

            if (markParam != null && markParam?.AsString() != string.Empty)
                return markParam.AsString();

            return string.Empty;
        }

        private int GetCountType(Element elementType)
        {
            var countTypeParam = elementType.LookupParameter(_rebarCountTypeParamName);

            if (countTypeParam != null)
                return countTypeParam.AsInteger();

            return 0;
        }

        private string GetMark(Element elementType)
        {
            // Маркировка типоразмера
            var markSystemTypeParam = elementType.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_MARK);

            // _Марка
            if (_rebarMarkParamGuid == Guid.Empty)
                _rebarMarkParamGuid = elementType.LookupParameter(_rebarMarkParamName) == null? Guid.Empty : elementType.LookupParameter(_rebarMarkParamName).GUID;

            var markSharedTypeParam = elementType.get_Parameter(_rebarMarkParamGuid);

            // _Наименование
            if (_rebarNominationParamGuid == Guid.Empty)
                _rebarNominationParamGuid = elementType.LookupParameter(_rebarNominationParamName).GUID;
            var markSharedNominationParam = elementType.get_Parameter(_rebarNominationParamGuid);

            if (markSharedTypeParam != null && markSharedTypeParam?.AsString() != string.Empty)
                return markSharedTypeParam.AsString();

            if (markSharedNominationParam != null && markSharedNominationParam?.AsString() != string.Empty)
                return markSharedNominationParam.AsString();

            if (markSystemTypeParam != null && markSystemTypeParam?.AsString() != string.Empty)
                return markSystemTypeParam.AsString();

            return string.Empty;
        }

        private string GetShapeImagePath(Element element)
        {
            var shapeImageSystemParam = element.get_Parameter(BuiltInParameter.ALL_MODEL_IMAGE);

            if (shapeImageSystemParam == null)
            {
                return null;
            }

            var shapeImageType = _doc.GetElement(shapeImageSystemParam.AsElementId()) as ImageType;
            if (shapeImageType == null)
            {
                return null;
            }

            ShapeImagesData.AddImageType(shapeImageType);
            return shapeImageType.Path;
        }

        private string GetConstructionType(Element element)
        {
            if (_rebarTypeOfConstructionParamGuid == Guid.Empty) 
                _rebarTypeOfConstructionParamGuid = element.LookupParameter(_rebarTypeOfConstructionParamName).GUID;
            var typeOfConstructionParamValue = element.get_Parameter(_rebarTypeOfConstructionParamGuid)?.AsString();

            var result = "(нет)";

            if (typeOfConstructionParamValue != null && typeOfConstructionParamValue != string.Empty)
                return typeOfConstructionParamValue;          

            return result;
        }

        private ConstructionType GetConstructionTypeEnum(string constructionType)
        {
            switch (constructionType)
            {
                case "Балки": return ConstructionType.Beam;
                case "Колонны": return ConstructionType.Column;
                case "Перекрытия": return ConstructionType.Floor;
                case "Лестницы": return ConstructionType.Stairs;
                case "Стены": return ConstructionType.Wall;
                case "Фундаменты": return ConstructionType.Foundation;
                default: return ConstructionType.Unknown;
            }
        }

        private string GetConstructionMark(Element element)
        {
            if (_rebarMarkOfConstructionParamGuid == Guid.Empty) 
                _rebarMarkOfConstructionParamGuid = element.LookupParameter(_rebarMarkOfConstructionParamName).GUID;
            var markOfConstructionParamValue = element.get_Parameter(_rebarMarkOfConstructionParamGuid)?.AsString();

            var result = "(нет)";

            if (markOfConstructionParamValue != null && markOfConstructionParamValue != string.Empty)
                return markOfConstructionParamValue;

            return result;
        }

        private int GetConstructionCount(Element element)
        {
            if (_rebarConstructionCountParamGuid == Guid.Empty) 
                _rebarConstructionCountParamGuid = element.LookupParameter(_rebarConstructionCountParamName).GUID;

            return element.get_Parameter(_rebarConstructionCountParamGuid).AsInteger();
        }

        private int GetTypicalFloor(Element element)
        {
            if (_rebarTypicalFloorParamGuid == Guid.Empty) 
                _rebarTypicalFloorParamGuid = element.LookupParameter(_rebarTypicalFloorParamName).GUID;

            return element.get_Parameter(_rebarTypicalFloorParamGuid).AsInteger();
        }

        private int GetTypicalFloorCount(Element element)
        {
            if (_rebarTypicalFloorCountParamGuid == Guid.Empty) 
                _rebarTypicalFloorCountParamGuid = element.LookupParameter(_rebarTypicalFloorCountParamName).GUID;

            return element.get_Parameter(_rebarTypicalFloorCountParamGuid).AsInteger();
        }

        private RebarLevel GetLevel(Element element)
        {
            if (_rebarLevelParamGuid == Guid.Empty)
                _rebarLevelParamGuid = element.LookupParameter(_rebarLevelParamName).GUID;
            var levelParamValue = element.get_Parameter(_rebarLevelParamGuid)?.AsString();

            if (levelParamValue != null && levelParamValue != string.Empty)
            {
                var elevation = GetLevelElevation(levelParamValue);
                return new RebarLevel(levelParamValue, elevation);
            }
                
            return new RebarLevel("(нет)", Double.MinValue);
        }

        private double GetLevelElevation(string levelParamValue)
        {
            var level = _levels.FirstOrDefault(l => l.LookupParameter(_rebarLevelParamName)?.AsString() == levelParamValue);
            var levelElevation = (level as Level).Elevation;
            return levelElevation;
        }

        private IList<Element> GetSystemLevels()
        {
            return new FilteredElementCollector(_doc)
                .OfCategory(BuiltInCategory.OST_Levels)
                .WhereElementIsNotElementType()
                .ToElements()
                .ToList();
        }

        private string GetSection(Element element)
        {
            if (_rebarSectionParamGuid == Guid.Empty)
            {
                if (element.LookupParameter(_rebarSectionParamName) != null)
                    _rebarSectionParamGuid = element.LookupParameter(_rebarSectionParamName).GUID;
            }
            var sectionParamValue = element.get_Parameter(_rebarSectionParamGuid)?.AsString();

            var result = "(нет)";

            if (sectionParamValue != null && sectionParamValue != string.Empty)
                return sectionParamValue;

            return result;
        }

        private string GetDefinition(Element elementType)
        {
            if (_rebarDefinitionParamGuid == Guid.Empty)
                _rebarDefinitionParamGuid = elementType.LookupParameter(_rebarDefinitionParamName).GUID;
            var markDefinitionParam = elementType.get_Parameter(_rebarDefinitionParamGuid);

            if (markDefinitionParam != null)
                return markDefinitionParam.AsString();

            return string.Empty;
        }

        private string GetClass(Element element, Element elementType)
        {
            var classTypeParam = elementType.LookupParameter(_rebarClassParamName);
            var classElementParam = element.LookupParameter(_rebarClassParamName);

            if (classTypeParam != null)
                return classTypeParam.AsString();

            if (classElementParam != null)
                return classElementParam.AsString();

            return string.Empty;
        }

        private double GetBarMass(Element element, Element elementType, int countTypeParamValue)
        {
            var massTypeParam = elementType.LookupParameter(_rebarMassParamName);
            var massElementParam = element.LookupParameter(_rebarMassParamName);
            var massPerLengthTypeParam = elementType.LookupParameter(_rebarMassPerLengthParamName);
            var massPerLengthElementParam = element.LookupParameter(_rebarMassPerLengthParamName);

            if (countTypeParamValue == 2)
            {
                if (massPerLengthTypeParam != null)
                    return UnitUtils.ConvertFromInternalUnits(massPerLengthTypeParam.AsDouble(), massPerLengthTypeParam.DisplayUnitType);

                if (massPerLengthElementParam != null)
                    return UnitUtils.ConvertFromInternalUnits(massPerLengthElementParam.AsDouble(), massPerLengthElementParam.DisplayUnitType);

                return 0;
            }
            
            if (massTypeParam != null)
                return Math.Round(massTypeParam.AsDouble(), 2);

            if (massElementParam != null)
                return Math.Round(massElementParam.AsDouble(), 2);

            return 0;
        }

        private double GetMass(Element element, Element elementType)
        {
            var massTypeParamValue = elementType.LookupParameter(_rebarMassParamName)?.AsDouble();
            var massElementParamValue = element.LookupParameter(_rebarMassParamName)?.AsDouble();

            if (massTypeParamValue != null)
                return Math.Round((double)massTypeParamValue, 2);

            if (massElementParamValue != null)
                return Math.Round((double)massElementParamValue, 2);

            return 0;
        }

        private double GetBarLengthСoefficient(Element elementType)
        {
            var lengthCoefficientParamValue = elementType.LookupParameter(_rebarLengthСoefficientParamName)?.AsDouble();

            if (lengthCoefficientParamValue != null)
                return (double)lengthCoefficientParamValue;

            return 1;
        }

        private double GetBarLength(Element element, int countTypeParamValue)
        {
            if (_rebarLengthParamGuid == Guid.Empty)
                _rebarLengthParamGuid = element.LookupParameter(_rebarLengthParamName).GUID;
            var lengthParam = element.get_Parameter(_rebarLengthParamGuid);

            if (countTypeParamValue == 2)
            {
                return 1;
            }

            if (lengthParam != null)
                return UnitUtils.ConvertFromInternalUnits(lengthParam.AsDouble(), lengthParam.DisplayUnitType);

            return 0;
        }

        private double GetLengthParamValue(Element element)
        {
            if (_rebarLengthParamGuid == Guid.Empty) 
                _rebarLengthParamGuid = element.LookupParameter(_rebarLengthParamName).GUID; 
            var lengthParam = element.get_Parameter(_rebarLengthParamGuid);

            if (lengthParam != null)
                return UnitUtils.ConvertFromInternalUnits(lengthParam.AsDouble(), lengthParam.DisplayUnitType);

            return 0;
        }

        private double GetCount(Element element, Element elementType, int countTypeParamValue, double barLengthСoefficient)
        {
            var countTypeSharedParam = elementType.LookupParameter(_rebarCountParamName);
            var countElementSharedParam = element.LookupParameter(_rebarCountParamName);
            var countSystemParam = element.get_Parameter(BuiltInParameter.REBAR_ELEM_QUANTITY_OF_BARS);

            if (countTypeParamValue == 2)
            {
                if (_rebarLengthParamGuid == Guid.Empty)
                    _rebarLengthParamGuid = element.LookupParameter(_rebarLengthParamName).GUID;
                var lengthParam = element.get_Parameter(_rebarLengthParamGuid);

                if (lengthParam != null)
                {
                    int barCountParamValue = 1;

                    if (countSystemParam != null)
                        barCountParamValue = countSystemParam.AsInteger();

                    else if (countElementSharedParam != null)
                        barCountParamValue = countElementSharedParam.AsInteger();

                    else if (countTypeSharedParam != null)
                        barCountParamValue = countTypeSharedParam.AsInteger();

                    return UnitUtils.ConvertFromInternalUnits(lengthParam.AsDouble() / 1000, lengthParam.DisplayUnitType) * barCountParamValue * barLengthСoefficient;
                }

                return 0;
            }

            if (countSystemParam != null)
                return countSystemParam.AsInteger();

            if (countElementSharedParam != null)
                return countElementSharedParam.AsInteger();

            if (countTypeSharedParam != null)
                return countTypeSharedParam.AsInteger();

            return 0;
        }

        private double GetDiameter(Element element, Element elementType)
        {
            var diameterTypeParam = elementType.LookupParameter(_rebarDiameterParamName);
            var diameterElementParam = element.LookupParameter(_rebarDiameterParamName);

            if (diameterTypeParam != null)
                return UnitUtils.ConvertFromInternalUnits(diameterTypeParam.AsDouble(), diameterTypeParam.DisplayUnitType);

            if (diameterElementParam != null)
                return UnitUtils.ConvertFromInternalUnits(diameterElementParam.AsDouble(), diameterElementParam.DisplayUnitType);

            return 0;
        }

        private void SetBarSizeParameters(ref Bar bar, Element barElement)
        {
            var previousSize = _sizes.First();

            foreach (var size in _sizes)
            {
                var sizeParam = barElement.LookupParameter(size.Key.Name);
                var previousSizeParam = barElement.LookupParameter(previousSize.Key.Name);

                if (sizeParam != null)
                {
                    if (previousSizeParam != null && previousSize.Key.Id == size.Key.Id)
                    {
                        var tempSizeMax = Math.Max(sizeParam.AsDouble(), previousSizeParam.AsDouble());

                        bar.Sizes[size.Key] = 0;
                        bar.Sizes[previousSize.Key] = UnitUtils.ConvertFromInternalUnits(tempSizeMax, sizeParam.DisplayUnitType);
                    }
                    else if (previousSize.Key.Id == size.Key.Id)
                    {
                        bar.Sizes[size.Key] = 0;
                        bar.Sizes[previousSize.Key] = UnitUtils.ConvertFromInternalUnits(sizeParam.AsDouble(), sizeParam.DisplayUnitType);
                    }
                    else
                    {
                        bar.Sizes[size.Key] = UnitUtils.ConvertFromInternalUnits(sizeParam.AsDouble(), sizeParam.DisplayUnitType);
                    }
                }
                else
                {
                    bar.Sizes[size.Key] = 0;
                }

                previousSize = size;
            }
        }

        private Element GetElementType(Element element)
        {
            var elementTypeId = element.GetTypeId();
            return _doc.GetElement(elementTypeId);
        }

        private void AddAllRebarsOfAssemblyToRebarAssembly(RebarAssembly rebarAssembly, ICollection<ElementId> allRebarIdsOfAsssembly)
        {
            foreach (var rebarId in allRebarIdsOfAsssembly)
            {
                var rebarElement = _doc.GetElement(rebarId);
                var elementType = GetElementType(rebarElement);

                if (rebarElement.Category.Name == "Несущая арматура")
                {
                    var rebarClassParamValue = GetClass(rebarElement, elementType);
                    var rebarDiameterParamValue = GetDiameter(rebarElement, elementType);
                    var barCountTypeParamValue = 1;
                    var rebarMassParamValue = GetMass(rebarElement, elementType);
                    var rebarShapeParamValue = rebarElement.LookupParameter(_rebarShapeParamName).AsString();

                    var rebar = new Rebar(rebarClassParamValue, rebarDiameterParamValue, rebarMassParamValue, rebarShapeParamValue)
                    {
                        Length = GetLengthParamValue(rebarElement),
                        Count = GetCount(rebarElement, elementType, barCountTypeParamValue, 1),
                        TypeOfAssembly = rebarAssembly.Type,
                        MarkOfAssembly = rebarAssembly.Mark
                    };

                    rebarAssembly.AddRebar(rebar);
                }
            }
        }

        private string GetCountTypeInfo(int countType)
        {
            if (countType == 1)
                return "шт.";

            if (countType == 2)
                return "м.п.";

            return "(не задан)";
        }
        #endregion МЕТОДЫ ПОЛУЧЕНИЯ ЗНАЧЕНИЙ ПАРАМЕТРОВ

        private void ExportToExcel()
        {
            ProjectData.FileName = _doc.Title + "_Арматура.xlsx";
            var filteredBars = BarsCollectionView.Cast<Bar>().ToList();

            var filteredRebarAssemblies = RebarAssembliesCollectionView.Cast<RebarAssembly>()
                .Where(rebarAssembly => rebarAssembly.Type.Contains("Каркас") || rebarAssembly.Type.Contains("Сетка"))
                .ToList();

            if (filteredBars.Any() || filteredRebarAssemblies.Any())
            {
                ProjectData.ProjectCode = _doc.ProjectInformation.Number;
                ProjectData.ProjectName = _doc.ProjectInformation.LookupParameter(_projectNameParamName)?.AsString();
                ProjectData.BuildingName = _doc.ProjectInformation.BuildingName;

                var tableManager = new TableManager(filteredBars, filteredRebarAssemblies);
                tableManager.CreateTable();
            }
            else
            {
                WarningWindow errorWindow = new WarningWindow("ПРЕДУПРЕЖДЕНИЕ", "Данные для экспорта отсутствуют");
                errorWindow.ShowDialog();
            }
        }
    }
}