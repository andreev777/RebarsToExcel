using Autodesk.Revit.DB;
using RebarsToExcel.Views;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace RebarsToExcel.Models.Bars
{
    /// <summary>
    /// Хранилище всех эскизов деталей.
    /// </summary>
    public static class ShapeImagesData
    {
        private static List<ImageType> _data = new List<ImageType>();
        private static string _path = "C:\\RebarScetch";

        /// <summary>
        /// Добавить эскиз детали в хранилище.
        /// </summary>
        public static void AddImageType(ImageType imageType)
        {
            if (_data.Any(existedImageType => existedImageType.Name == imageType.Name))
                return;

            _data.Add(imageType);
        }

        /// <summary>
        /// Сохранить эскизы деталей в отдельую директорию на компьютере.
        /// </summary>
        public static void SaveToFolder()
        {
            if (!_data.Any())
                return;

            if (!Directory.Exists(_path))
            {
                Directory.CreateDirectory(_path);
            }
            else
            {
                try
                {
                    DeleteImagesFolder();
                    Directory.CreateDirectory(_path);
                }
                catch (Exception e)
                {
                    var exceptionWindow = new ExceptionWindow(e.Message, e.StackTrace);
                    exceptionWindow.ShowDialog();
                    return;
                }
            }

            foreach (var imageType in _data)
            {
                using (var image = imageType.GetImage())
                {
                    image.Save(Path.Combine(_path, imageType.Name));
                }
            }
        }

        private static void DeleteImagesFolder()
        {
            var directoryInfo = new DirectoryInfo(_path);
            foreach (var file in directoryInfo.GetFiles())
            {
                file.Delete();
            }

            foreach (var directory in directoryInfo.GetDirectories())
            {
                directory.Delete(true);
            }

            Directory.Delete(_path);
        }
    }
}