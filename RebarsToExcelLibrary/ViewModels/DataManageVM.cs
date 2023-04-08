using Autodesk.Revit.DB;
using Prism.Mvvm;
using RebarsToExcel.Commands;
using RebarsToExcel.Models;
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
        private readonly string beamGroupModelParamValue = "Балки монолитные";
        private readonly string columnGroupModelParamValue = "Колонны монолитные";
        private readonly string floorGroupModelParamValue = "Перекрытия монолитные";
        private readonly string wallGroupModelParamValue = "Стены монолитные";
        private readonly string rebarGroupModelParamValue = "Детали";
        private readonly string assemblyGroupModelParamValue = "Сборочные единицы";

        private readonly string rebarClassParamName = "_Класс арматуры";
        private readonly string rebarDiameterParamName = "_Диаметр стержня";
        private readonly string rebarMassParamName = "_Масса";
        private readonly string rebarShapeParamName = "_Форма стержня";
        private readonly string rebarLengthParamName = "_Длина стержня";
        private readonly string rebarLengthСoefficientParamName = "_Коэф. перепуска";
        private readonly string rebarCountTypeParamName = "_Тип подсчёта количества";
        private readonly string rebarCountParamName = "_Количество";
        private readonly string rebarMarkParamName = "_Марка";
        private readonly string rebarDefinitionParamName = "_Наименование";
        private readonly string rebarTypeOfConstructionParamName = "_Тип основы";
        private readonly string rebarMarkOfConstructionParamName = "_Метка основы";
        private readonly string rebarConstructionCountParamName = "_Количество основ";
        private readonly string rebarTypicalFloorParamName = "_Типовой этаж";
        private readonly string rebarTypicalFloorCountParamName = "_Количество типовых этажей";
        private readonly string rebarLevelParamName = "_Этаж";
        private readonly string rebarSectionParamName = "_Секция";

        private Guid rebarLengthParamGuid = Guid.Empty;
        private Guid rebarMarkParamGuid = Guid.Empty;
        private Guid rebarDefinitionParamGuid = Guid.Empty;
        private Guid rebarTypeOfConstructionParamGuid = Guid.Empty;
        private Guid rebarMarkOfConstructionParamGuid = Guid.Empty;
        private Guid rebarConstructionCountParamGuid = Guid.Empty;
        private Guid rebarTypicalFloorParamGuid = Guid.Empty;
        private Guid rebarTypicalFloorCountParamGuid = Guid.Empty;
        private Guid rebarLevelParamGuid = Guid.Empty;
        private Guid rebarSectionParamGuid = Guid.Empty;

        private Document _doc;
        private List<Element> _levels;

        public bool IsDataEmpty { get; private set; } = false;
        public string DocumentTitle { get; private set; }
        public double BarsTotalCount { get; set; }
        public double RebarAssembliesTotalCount { get; set; }
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

        public List<string> BarSections { get; set; }
        public List<string> BarConstructionTypes { get; set; }

        private List<string> _barConstructionMarks;
        public List<string> BarConstructionMarks
        {
            get => _barConstructionMarks;
            set
            {
                _barConstructionMarks = value;
                RaisePropertyChanged(nameof(BarConstructionMarks));
            }
        }

        public List<RebarLevel> BarLevels { get; set; }
        public ICollectionView BarsCollectionView { get; set; }
        public Bar SelectedBar { get; set; }
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
        
        public List<string> RebarAssemblySections { get; set; }
        public List<string> RebarAssemblyConstructionTypes { get; set; }

        private List<string> _rebarAssemblyConstructionMarks;
        public List<string> RebarAssemblyConstructionMarks 
        { 
            get => _rebarAssemblyConstructionMarks; 
            set
            {
                _rebarAssemblyConstructionMarks = value;
                RaisePropertyChanged(nameof(RebarAssemblyConstructionMarks));
            }
        }

        public List<RebarLevel> RebarAssemblyLevels { get; set; }
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

            BackgroundWorker.DoWork += (object sender, DoWorkEventArgs e) =>
            {
                AnalyzeAllBars(allBarElements, typicalFloors);
                AnalyzeAllRebarAssemblies(allGenericModelElements, allSystemAssemblyElements, typicalFloors);
            };

            BackgroundWorker.RunWorkerCompleted += (object sender, RunWorkerCompletedEventArgs e) =>
            {
                if (analyzeWindow != null)
                {
                    analyzeWindow.Close();
                }

                // Получение данных деталей
                var barsData = BarsData.GetData();
                BarsCollectionView = CollectionViewSource.GetDefaultView(barsData);

                BarLevels = BarsData.GetLevels();
                BarSections = BarsData.GetSections();
                BarConstructionTypes = BarsData.GetConstructionTypes();

                SelectedBarSection = BarSections.FirstOrDefault();
                SelectedBarConstructionType = BarConstructionTypes.FirstOrDefault();
                SelectedBarConstructionMark = "(все)";

                // Получение данных сборочных единиц

                var rebarAssembliesData = RebarAssembliesData.GetData();
                RebarAssembliesCollectionView = CollectionViewSource.GetDefaultView(rebarAssembliesData);

                RebarAssemblyLevels = RebarAssembliesData.GetLevels();
                RebarAssemblySections = RebarAssembliesData.GetSections();
                RebarAssemblySections.Insert(0, "(все)");
                RebarAssemblyConstructionTypes = RebarAssembliesData.GetConstructionTypes();
                RebarAssemblyConstructionTypes.Insert(0, "(все)");

                SelectedRebarAssemblySection = RebarAssemblySections.FirstOrDefault();
                SelectedRebarAssemblyConstructionType = RebarAssemblyConstructionTypes.FirstOrDefault();
                SelectedRebarAssemblyConstructionMark = "(все)";

                FilterRebarAssemblyElements();
            };

            analyzeWindow.Show();

            BackgroundWorker.RunWorkerAsync();
        }

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
            return constructions.Where(x => GetElementType(x).get_Parameter(BuiltInParameter.ALL_MODEL_MODEL)?.AsString() == beamGroupModelParamValue).ToList();
        }

        private List<Element> GetAllColumns(List<Element> constructions)
        {
            return constructions.Where(x => GetElementType(x).get_Parameter(BuiltInParameter.ALL_MODEL_MODEL)?.AsString() == columnGroupModelParamValue).ToList();
        }

        private List<Element> GetAllFloors(List<Element> constructions)
        {
            return constructions.Where(x => GetElementType(x).get_Parameter(BuiltInParameter.ALL_MODEL_MODEL)?.AsString() == floorGroupModelParamValue).ToList();
        }

        private List<Element> GetAllWalls(List<Element> constructions)
        {
            return constructions.Where(x => GetElementType(x).get_Parameter(BuiltInParameter.ALL_MODEL_MODEL)?.AsString() == wallGroupModelParamValue).ToList();
        }
        #endregion

        #region ПОЛУЧЕНИЕ ТИПОВЫХ ЭТАЖЕЙ
        private List<TypicalFloor> GetAllTypicalFloors()
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

        private List<TypicalFloor> GetAllBeamTypicalFloors(List<Element> beams)
        {
            var beamsWithMultipleTypicalFloors = beams.Where(element => GetTypicalFloorCountParamValue(element) > 1).ToList();

            var beamTypicalLevels = beamsWithMultipleTypicalFloors.GroupBy(element => GetTypicalFloorParamValue(element))
                .Select(element => new TypicalFloor(ConstructionType.Beam, element.Key, element.Select(el => GetLevelParamValue(el)).Distinct(new LevelComparer()).ToList()))
                .ToList();

            return beamTypicalLevels;
        }

        private List<TypicalFloor> GetAllColumnTypicalFloors(List<Element> columns)
        {
            var columnsWithMultipleTypicalFloors = columns.Where(element => GetTypicalFloorCountParamValue(element) > 1).ToList();

            var columnTypicalLevels = columnsWithMultipleTypicalFloors.GroupBy(element => GetTypicalFloorParamValue(element))
                .Select(element => new TypicalFloor(ConstructionType.Column, element.Key, element.Select(el => GetLevelParamValue(el)).Distinct(new LevelComparer()).ToList()))
                .ToList();

            return columnTypicalLevels;
        }

        private List<TypicalFloor> GetAllFloorTypicalFloors(List<Element> floors)
        {
            var floorsWithMultipleTypicalFloors = floors.Where(element => GetTypicalFloorCountParamValue(element) > 1).ToList();

            var floorTypicalLevels = floorsWithMultipleTypicalFloors.GroupBy(element => GetTypicalFloorParamValue(element))
                .Select(element => new TypicalFloor(ConstructionType.Floor, element.Key, element.Select(el => GetLevelParamValue(el)).Distinct(new LevelComparer()).ToList()))
                .ToList();

            return floorTypicalLevels;
        }

        private List<TypicalFloor> GetAllWallTypicalFloors(List<Element> walls)
        {
            var wallsWithMultipleTypicalFloors = walls.Where(element => GetTypicalFloorCountParamValue(element) > 1).ToList();

            var wallTypicalLevels = wallsWithMultipleTypicalFloors.GroupBy(element => GetTypicalFloorParamValue(element))
                .Select(element => new TypicalFloor(ConstructionType.Wall, element.Key, element.Select(el => GetLevelParamValue(el)).Distinct(new LevelComparer()).ToList()))
                .ToList();

            return wallTypicalLevels;
        }
        #endregion

        #region ПОЛУЧЕНИЕ И АНАЛИЗ АРМАТУРЫ
        private List<Element> GetAllBars()
        {
            ParameterValueProvider groupModelProvider = new ParameterValueProvider(new ElementId(BuiltInParameter.ALL_MODEL_MODEL));
            FilterStringRuleEvaluator groupModelEvaluator = new FilterStringEquals();
            FilterRule groupModelFilterRule = new FilterStringRule(groupModelProvider, groupModelEvaluator, rebarGroupModelParamValue, false);
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

        private List<Element> GetAllGenericModelElements()
        {
            ParameterValueProvider provider = new ParameterValueProvider(new ElementId(BuiltInParameter.ALL_MODEL_MODEL));
            FilterStringRuleEvaluator evaluator = new FilterStringEquals();
            FilterRule groupModelFilterRule = new FilterStringRule(provider, evaluator, assemblyGroupModelParamValue, false);
            ElementParameterFilter groupModelFilter = new ElementParameterFilter(groupModelFilterRule);

            return new FilteredElementCollector(_doc)
                .OfCategory(BuiltInCategory.OST_GenericModel)
                .WhereElementIsNotElementType()
                .WherePasses(groupModelFilter)
                .ToElements()
                .ToList();
        }

        private List<Element> GetAllSystemAssemblyElements()
        {
            ParameterValueProvider provider = new ParameterValueProvider(new ElementId(BuiltInParameter.ALL_MODEL_MODEL));
            FilterStringRuleEvaluator evaluator = new FilterStringEquals();
            FilterRule groupModelFilterRule = new FilterStringRule(provider, evaluator, assemblyGroupModelParamValue, false);
            ElementParameterFilter groupModelFilter = new ElementParameterFilter(groupModelFilterRule);

            return new FilteredElementCollector(_doc)
                .OfCategory(BuiltInCategory.OST_Assemblies)
                .WhereElementIsNotElementType()
                .WherePasses(groupModelFilter)
                .ToElements()
                .ToList();
        }

        private void AnalyzeAllBars(IList<Element> allBarElements, List<TypicalFloor> typicalFloors)
        {
            foreach (var barElement in allBarElements) //Перебираем всю арматуру
            {
                var elementType = GetElementType(barElement);

                var barClassParamValue = GetClassParamValue(barElement, elementType); //_Класс
                var barDiameterParamValue = GetDiameterParamValue(barElement, elementType); //_Диаметр
                var barCountParamValue = GetCountParamValue(barElement, elementType); //_Количество
                var barConstructionTypeParamValue = GetConstructionTypeParamValue(barElement); //_Тип основы
                var barCountTypeParamValue = GetCountTypeParamValue(elementType); //_Тип подсчета количества
                var barLengthСoefficient = GetBarLengthСoefficient(elementType); //_Коэф. перепуска
                var barLengthParamValue = GetBarLengthParamValue(barElement, barCountTypeParamValue, barCountParamValue, barLengthСoefficient); //_Длина
                var barMassParamValue = GetBarMassParamValue(barElement, elementType, barCountTypeParamValue, barCountParamValue); //_Масса
                var barShapeParamValue = barElement.LookupParameter(rebarShapeParamName).AsString(); //_Номер формы

                var bar = new Bar(barClassParamValue, barDiameterParamValue, barMassParamValue, barShapeParamValue)
                {
                    Id = barElement.Id,
                    Position = GetPositionParamValue(barElement), //Марка
                    Length = barLengthParamValue,
                    CountType = barCountTypeParamValue,
                    CountTypeInfo = GetCountTypeInfo(barCountTypeParamValue),
                    Count = barCountParamValue,
                    Level = GetLevelParamValue(barElement),
                    Section = GetSectionParamValue(barElement), //_Секция
                    ConstructionType = barConstructionTypeParamValue,
                    ConstructionTypeEnum = GetConstructionTypeEnum(barConstructionTypeParamValue),
                    ConstructionMark = GetConstructionMarkParamValue(barElement), //_Метка основы
                    ConstructionCount = GetConstructionCountParamValue(barElement), //_Количество основ
                    TypicalFloor = GetTypicalFloorParamValue(barElement), //_Типовой этаж
                    TypicalFloorCount = GetTypicalFloorCountParamValue(barElement), //_Количество типовых этажей
                };

                BarsData.AddBar(bar);
                BarsProgressCounter++;
            }

            BarsData.AnalyzeDataByConstructionCount();
            BarsData.AnalyzeDataByTypicalFloorCount(typicalFloors);
        }

        private void AnalyzeAllRebarAssemblies(IList<Element> allGenericModelElements, IList<Element> allSystemAssemblyElements, List<TypicalFloor> typicalFloors)
        {
            foreach (var genericModelElement in allGenericModelElements) //Перебираем все Обобщенные модели
            {
                var elementType = GetElementType(genericModelElement);
                var descriptionTypeParamValue = GetDescriptionParamValue(elementType); //Описание
                var globalModelTypeParamValue = GetGroupModelParamValue(elementType); //Группа модели

                var markParamValue = GetMarkParamValue(elementType); //Маркировка типоразмера, _Марка или _Наименование
                var massParamValue = GetMassParamValue(genericModelElement, elementType); //_Масса
                var constructionTypeParamValue = GetConstructionTypeParamValue(genericModelElement); //_Тип основы

                var rebarAssembly = new RebarAssembly(descriptionTypeParamValue, markParamValue, globalModelTypeParamValue, massParamValue)
                {
                    Id = genericModelElement.Id,
                    Definition = GetDefinitionParamValue(genericModelElement), //_Обозначение
                    ConstructionType = constructionTypeParamValue,
                    ConstructionTypeEnum = GetConstructionTypeEnum(constructionTypeParamValue),
                    ConstructionMark = GetConstructionMarkParamValue(genericModelElement), //_Метка основы
                    ConstructionCount = GetConstructionCountParamValue(genericModelElement), //_Количество основ
                    TypicalFloor = GetTypicalFloorParamValue(genericModelElement), //_Типовой этаж
                    TypicalFloorCount = GetTypicalFloorCountParamValue(genericModelElement), //_Количество типовых этажей
                    Level = GetLevelParamValue(genericModelElement), //_Этаж
                    Section = GetSectionParamValue(genericModelElement) //_Секция
                };

                var allRebarIdsOfAsssembly = (genericModelElement as FamilyInstance).GetSubComponentIds();
                AddAllRebarsOfAssemblyToRebarAssembly(elementType, rebarAssembly, allRebarIdsOfAsssembly); //Добавляем всю вложенную арматуру семейства в rebarAssembly

                RebarAssembliesData.AddRebarAssembly(rebarAssembly);
                RebarAssembliesProgressCounter++;
            }

            foreach (var assemblyElement in allSystemAssemblyElements) //Перебираем все Сборки
            {
                var elementType = GetElementType(assemblyElement);
                var descriptionTypeParamValue = GetDescriptionParamValue(elementType); //Описание
                var globalModelTypeParamValue = GetGroupModelParamValue(elementType); //Группа модели

                var markParamValue = GetMarkParamValue(elementType); //Маркировка типоразмера или _Марка
                var massParamValue = GetMassParamValue(assemblyElement, elementType); //_Масса
                var constructionTypeParamValue = GetConstructionTypeParamValue(assemblyElement); //_Тип основы

                var rebarAssembly = new RebarAssembly(descriptionTypeParamValue, markParamValue, globalModelTypeParamValue, massParamValue)
                {
                    Id = assemblyElement.Id,
                    Definition = GetDefinitionParamValue(assemblyElement), //_Обозначение
                    ConstructionType = constructionTypeParamValue,
                    ConstructionTypeEnum = GetConstructionTypeEnum(constructionTypeParamValue),
                    ConstructionMark = GetConstructionMarkParamValue(assemblyElement), //_Метка основы
                    ConstructionCount = GetConstructionCountParamValue(assemblyElement), //_Количество основ
                    TypicalFloor = GetTypicalFloorParamValue(assemblyElement), //_Типовой этаж
                    TypicalFloorCount = GetTypicalFloorCountParamValue(assemblyElement), //_Количество типовых этажей
                    Level = GetLevelParamValue(assemblyElement), //_Этаж
                    Section = GetSectionParamValue(assemblyElement) //_Секция
                };

                var allRebarIdsOfAsssembly = (assemblyElement as AssemblyInstance).GetMemberIds();
                AddAllRebarsOfAssemblyToRebarAssembly(elementType, rebarAssembly, allRebarIdsOfAsssembly); //Добавляем всю вложенную арматуру семейства в rebarAssembly

                RebarAssembliesData.AddRebarAssembly(rebarAssembly);
                RebarAssembliesProgressCounter++;
            }

            //Анализ сборочных единиц по количеству основ и типовым этажам
            RebarAssembliesData.AnalyzeDataByConstructionCount();
            RebarAssembliesData.AnalyzeDataByTypicalFloorCount(typicalFloors);
        }
        #endregion

        #region ПОЛУЧЕНИЕ ПАРАМЕТРОВ
        private string GetDescriptionParamValue(Element elementType)
        {
            return elementType.get_Parameter(BuiltInParameter.ALL_MODEL_DESCRIPTION)?.AsString();
        }

        private string GetGroupModelParamValue(Element elementType)
        {
            return elementType.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL)?.AsString();
        }

        private string GetPositionParamValue(Element element)
        {
            var markParam = element.get_Parameter(BuiltInParameter.ALL_MODEL_MARK);

            if (markParam != null && markParam?.AsString() != string.Empty)
                return markParam.AsString();

            return string.Empty;
        }

        private int GetCountTypeParamValue(Element elementType)
        {
            var countTypeParam = elementType.LookupParameter(rebarCountTypeParamName);

            if (countTypeParam != null)
                return countTypeParam.AsInteger();

            return 0;
        }

        private string GetMarkParamValue(Element elementType)
        {
            // Маркировка типоразмера
            var markSystemTypeParam = elementType.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_MARK);

            // _Марка
            if (rebarMarkParamGuid == Guid.Empty)
                rebarMarkParamGuid = elementType.LookupParameter(rebarMarkParamName) == null? Guid.Empty : elementType.LookupParameter(rebarMarkParamName).GUID;

            var markSharedTypeParam = elementType.get_Parameter(rebarMarkParamGuid);

            // _Наименование
            if (rebarDefinitionParamGuid == Guid.Empty)
                rebarDefinitionParamGuid = elementType.LookupParameter(rebarDefinitionParamName).GUID;
            var markSharedDefinitionParam = elementType.get_Parameter(rebarDefinitionParamGuid);

            if (markSharedTypeParam != null && markSharedTypeParam?.AsString() != string.Empty)
                return markSharedTypeParam.AsString();

            if (markSharedDefinitionParam != null && markSharedDefinitionParam?.AsString() != string.Empty)
                return markSharedDefinitionParam.AsString();

            if (markSystemTypeParam != null && markSystemTypeParam?.AsString() != string.Empty)
                return markSystemTypeParam.AsString();

            return string.Empty;
        }

        private string GetConstructionTypeParamValue(Element element)
        {
            if (rebarTypeOfConstructionParamGuid == Guid.Empty) 
                rebarTypeOfConstructionParamGuid = element.LookupParameter(rebarTypeOfConstructionParamName).GUID;
            var typeOfConstructionParamValue = element.get_Parameter(rebarTypeOfConstructionParamGuid)?.AsString();

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

        private string GetConstructionMarkParamValue(Element element)
        {
            if (rebarMarkOfConstructionParamGuid == Guid.Empty) 
                rebarMarkOfConstructionParamGuid = element.LookupParameter(rebarMarkOfConstructionParamName).GUID;
            var markOfConstructionParamValue = element.get_Parameter(rebarMarkOfConstructionParamGuid)?.AsString();

            var result = "(нет)";

            if (markOfConstructionParamValue != null && markOfConstructionParamValue != string.Empty)
                return markOfConstructionParamValue;

            return result;
        }

        private int GetConstructionCountParamValue(Element element)
        {
            if (rebarConstructionCountParamGuid == Guid.Empty) 
                rebarConstructionCountParamGuid = element.LookupParameter(rebarConstructionCountParamName).GUID;

            return element.get_Parameter(rebarConstructionCountParamGuid).AsInteger();
        }

        private int GetTypicalFloorParamValue(Element element)
        {
            if (rebarTypicalFloorParamGuid == Guid.Empty) 
                rebarTypicalFloorParamGuid = element.LookupParameter(rebarTypicalFloorParamName).GUID;

            return element.get_Parameter(rebarTypicalFloorParamGuid).AsInteger();
        }

        private int GetTypicalFloorCountParamValue(Element element)
        {
            if (rebarTypicalFloorCountParamGuid == Guid.Empty) 
                rebarTypicalFloorCountParamGuid = element.LookupParameter(rebarTypicalFloorCountParamName).GUID;

            return element.get_Parameter(rebarTypicalFloorCountParamGuid).AsInteger();
        }

        private RebarLevel GetLevelParamValue(Element element)
        {
            if (rebarLevelParamGuid == Guid.Empty)
                rebarLevelParamGuid = element.LookupParameter(rebarLevelParamName).GUID;
            var levelParamValue = element.get_Parameter(rebarLevelParamGuid)?.AsString();

            if (levelParamValue != null && levelParamValue != string.Empty)
            {
                var elevation = GetLevelElevation(levelParamValue);
                return new RebarLevel(levelParamValue, elevation);
            }
                
            return new RebarLevel("(нет)", Double.MinValue);
        }

        private double GetLevelElevation(string levelParamValue)
        {
            var level = _levels.FirstOrDefault(l => l.LookupParameter(rebarLevelParamName)?.AsString() == levelParamValue);
            var levelElevation = (level as Level).Elevation;
            return levelElevation;
        }

        private List<Element> GetSystemLevels()
        {
            return new FilteredElementCollector(_doc)
                .OfCategory(BuiltInCategory.OST_Levels)
                .WhereElementIsNotElementType()
                .ToElements()
                .ToList();
        }

        private string GetSectionParamValue(Element element)
        {
            if (rebarSectionParamGuid == Guid.Empty)
            {
                if (element.LookupParameter(rebarSectionParamName) != null)
                    rebarSectionParamGuid = element.LookupParameter(rebarSectionParamName).GUID;
            }
            var sectionParamValue = element.get_Parameter(rebarSectionParamGuid)?.AsString();

            var result = "(нет)";

            if (sectionParamValue != null && sectionParamValue != string.Empty)
                return sectionParamValue;

            return result;
        }

        private string GetDefinitionParamValue(Element elementType)
        {
            if (rebarDefinitionParamGuid == Guid.Empty) 
                rebarDefinitionParamGuid = elementType.LookupParameter(rebarDefinitionParamName).GUID;
            var markDefinitionParam = elementType.get_Parameter(rebarDefinitionParamGuid);

            if (markDefinitionParam != null)
                return markDefinitionParam.AsString();

            return string.Empty;
        }

        private string GetClassParamValue(Element element, Element elementType)
        {
            var classTypeParam = elementType.LookupParameter(rebarClassParamName);
            var classElementParam = element.LookupParameter(rebarClassParamName);

            if (classTypeParam != null)
                return classTypeParam.AsString();

            if (classElementParam != null)
                return classElementParam.AsString();

            return string.Empty;
        }

        private double GetBarMassParamValue(Element element, Element elementType, int countTypeParamValue, int barCountParamValue)
        {
            var massTypeParamValue = elementType.LookupParameter(rebarMassParamName)?.AsDouble();
            var massElementParamValue = element.LookupParameter(rebarMassParamName)?.AsDouble();

            if (countTypeParamValue == 2)
            {
                if (massTypeParamValue != null)
                    return barCountParamValue * Math.Round((double)massTypeParamValue, 2);

                if (massElementParamValue != null)
                    return barCountParamValue * Math.Round((double)massElementParamValue, 2);

                return 0;
            }
            
            if (massTypeParamValue != null)
                return Math.Round((double)massTypeParamValue, 2);

            if (massElementParamValue != null)
                return Math.Round((double)massElementParamValue, 2);

            return 0;
        }

        private double GetMassParamValue(Element element, Element elementType)
        {
            var massTypeParamValue = elementType.LookupParameter(rebarMassParamName)?.AsDouble();
            var massElementParamValue = element.LookupParameter(rebarMassParamName)?.AsDouble();

            if (massTypeParamValue != null)
                return Math.Round((double)massTypeParamValue, 2);

            if (massElementParamValue != null)
                return Math.Round((double)massElementParamValue, 2);

            return 0;
        }

        private double GetBarLengthСoefficient(Element elementType)
        {
            var lengthCoefficientParamValue = elementType.LookupParameter(rebarLengthСoefficientParamName)?.AsDouble();

            if (lengthCoefficientParamValue != null)
                return (double)lengthCoefficientParamValue;

            return 1;
        }

        private double GetBarLengthParamValue(Element element, int countTypeParamValue, int barCountParamValue, double barLengthСoefficient)
        {
            if (rebarLengthParamGuid == Guid.Empty)
                rebarLengthParamGuid = element.LookupParameter(rebarLengthParamName).GUID;
            var lengthParam = element.get_Parameter(rebarLengthParamGuid);

            if (countTypeParamValue == 2)
            {
                if (lengthParam != null)
                    return UnitUtils.ConvertFromInternalUnits(lengthParam.AsDouble(), lengthParam.DisplayUnitType) * barCountParamValue * barLengthСoefficient;

                return 0;
            }

            if (lengthParam != null)
                return UnitUtils.ConvertFromInternalUnits(lengthParam.AsDouble(), lengthParam.DisplayUnitType);

            return 0;
        }

        private double GetLengthParamValue(Element element)
        {
            if (rebarLengthParamGuid == Guid.Empty) 
                rebarLengthParamGuid = element.LookupParameter(rebarLengthParamName).GUID; 
            var lengthParam = element.get_Parameter(rebarLengthParamGuid);

            if (lengthParam != null)
                return UnitUtils.ConvertFromInternalUnits(lengthParam.AsDouble(), lengthParam.DisplayUnitType);

            return 0;
        }

        private int GetCountParamValue(Element element, Element elementType)
        {
            var countTypeSharedParam = elementType.LookupParameter(rebarCountParamName);
            var countElementSharedParam = element.LookupParameter(rebarCountParamName);
            var countSystemParam = element.get_Parameter(BuiltInParameter.REBAR_ELEM_QUANTITY_OF_BARS);

            if (countSystemParam != null)
                return countSystemParam.AsInteger();

            if (countElementSharedParam != null)
                return countElementSharedParam.AsInteger();

            if (countTypeSharedParam != null)
                return countTypeSharedParam.AsInteger();

            return 0;
        }

        private double GetDiameterParamValue(Element element, Element elementType)
        {
            var diameterTypeParam = elementType.LookupParameter(rebarDiameterParamName);
            var diameterElementParam = element.LookupParameter(rebarDiameterParamName);

            if (diameterTypeParam != null)
                return UnitUtils.ConvertFromInternalUnits(diameterTypeParam.AsDouble(), diameterTypeParam.DisplayUnitType);

            if (diameterElementParam != null)
                return UnitUtils.ConvertFromInternalUnits(diameterElementParam.AsDouble(), diameterElementParam.DisplayUnitType);

            return 0;
        }

        private Element GetElementType(Element element)
        {
            var elementTypeId = element.GetTypeId();
            return _doc.GetElement(elementTypeId);
        }

        private void AddAllRebarsOfAssemblyToRebarAssembly(Element elementType, RebarAssembly rebarAssembly, ICollection<ElementId> allRebarIdsOfAsssembly)
        {
            foreach (var rebarId in allRebarIdsOfAsssembly)
            {
                var rebarElement = _doc.GetElement(rebarId);

                if (rebarElement.Category.Name == "Несущая арматура")
                {
                    var rebarClassParamValue = GetClassParamValue(rebarElement, elementType);
                    var rebarDiameterParamValue = GetDiameterParamValue(rebarElement, elementType);
                    var rebarMassParamValue = GetMassParamValue(rebarElement, elementType);
                    var rebarShapeParamValue = rebarElement.LookupParameter(rebarShapeParamName).AsString();

                    var rebar = new Rebar(rebarClassParamValue, rebarDiameterParamValue, rebarMassParamValue, rebarShapeParamValue)
                    {
                        Length = GetLengthParamValue(rebarElement),
                        Count = GetCountParamValue(rebarElement, elementType),
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
                return "м.п";

            return "(не задан)";
        }
        #endregion МЕТОДЫ ПОЛУЧЕНИЯ ЗНАЧЕНИЙ ПАРАМЕТРОВ

        private void ExportToExcel()
        {
            string fileName = _doc.Title + "_Сборочные единицы.xlsx";
            var filteredRebarAssemblies = RebarAssembliesCollectionView.Cast<RebarAssembly>()
                .Where(rebarAssembly => rebarAssembly.Type.Contains("Каркас") || rebarAssembly.Type.Contains("Сетка"))
                .ToList();

            if (filteredRebarAssemblies.Any())
            {
                FileManager.Save(filteredRebarAssemblies, fileName);
            }
            else
            {
                WarningWindow errorWindow = new WarningWindow("ПРЕДУПРЕЖДЕНИЕ", "Данные для экспорта отсутствуют");
                errorWindow.ShowDialog();
            }


        }
    }
}