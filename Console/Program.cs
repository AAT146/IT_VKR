using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using MathNet.Numerics.Distributions;
using MathNet.Numerics.Statistics;
using MathNet.Numerics.LinearAlgebra;
using MathNet.Numerics.LinearAlgebra.Double;
using OfficeOpenXml;
using Microsoft.SqlServer.Server;

namespace ConsoleApplication
{
	public class MeasurementData
	{
		public DateTime Timestamp { get; set; }
		public double Value { get; set; }
	}

	public class ClusterResult
	{
		public int ClusterId { get; set; }
		public List<double> Values { get; set; }
		public string Distribution { get; set; }
		public Dictionary<string, double> Parameters { get; set; }
		public double Weight { get; set; }
	}

	public class DistributionAnalyzer
	{
		public static List<ClusterResult> AnalyzeWithKMeans(List<double> data, int clusterCount)
		{
			// Подготовка данных для кластеризации
			var dataPoints = data.Select(value => Vector<double>.Build.DenseOfArray(new[] { value })).ToList();

			// Инициализация центроидов
			var random = new Random();
			var centroids = dataPoints.OrderBy(_ => random.Next()).Take(clusterCount).ToList();

			// K-Means кластеризация
			var clusters = new Dictionary<int, List<double>>();
			for (int i = 0; i < clusterCount; i++) clusters[i] = new List<double>();

			bool changed;
			do
			{
				// Очистка кластеров
				foreach (var key in clusters.Keys) clusters[key].Clear();

				// Распределение данных по кластерам
				foreach (var point in dataPoints)
				{
					var nearestCluster = centroids
						.Select((centroid, index) => (Distance: (point - centroid).L2Norm(), Index: index))
						.OrderBy(x => x.Distance)
						.First()
						.Index;

					clusters[nearestCluster].Add(point[0]);
				}

				// Обновление центроидов
				changed = false;
				for (int i = 0; i < clusterCount; i++)
				{
					if (clusters[i].Count == 0) continue;

					var newCentroid = Vector<double>.Build.DenseOfArray(new[] { clusters[i].Average() });
					if (!centroids[i].Equals(newCentroid))
					{
						centroids[i] = newCentroid;
						changed = true;
					}
				}
			} while (changed);

			// Анализ распределений для каждого кластера
			var results = new List<ClusterResult>();
			foreach (var cluster in clusters.Where(c => c.Value.Any()))
			{
				var clusterValues = cluster.Value;
				var (distribution, parameters) = AnalyzeCluster(clusterValues);
				var weight = (double)clusterValues.Count / data.Count;

				results.Add(new ClusterResult
				{
					ClusterId = cluster.Key,
					Values = clusterValues,
					Distribution = distribution,
					Parameters = parameters,
					Weight = weight
				});
			}

			return results;
		}

		private static (string, Dictionary<string, double>) AnalyzeCluster(List<double> data)
		{
			// Параметры для нормального распределения
			var mean = Statistics.Mean(data);
			var stdDev = Statistics.StandardDeviation(data);

			var normalDist = new Normal(mean, stdDev);

			// Параметры для равномерного распределения
			var uniformDist = new ContinuousUniform(data.Min(), data.Max());

			// Параметры для экспоненциального распределения
			var lambda = 1.0 / mean;
			var exponentialDist = new Exponential(lambda);

			// Оценка качества соответствия
			double normalFit = ComputeFit(data, normalDist);
			double uniformFit = ComputeFit(data, uniformDist);
			double exponentialFit = ComputeFit(data, exponentialDist);

			// Выбор лучшего распределения
			if (normalFit < uniformFit && normalFit < exponentialFit)
			{
				return ("Normal", new Dictionary<string, double> { { "Mean", mean }, { "StdDev", stdDev } });
			}
			else if (uniformFit < normalFit && uniformFit < exponentialFit)
			{
				return ("Uniform", new Dictionary<string, double> { { "Min", data.Min() }, { "Max", data.Max() } });
			}
			else
			{
				return ("Exponential", new Dictionary<string, double> { { "Lambda", lambda } });
			}
		}

		private static double ComputeFit(List<double> data, IContinuousDistribution dist)
		{
			double logLikelihood = 0;
			foreach (var value in data)
			{
				logLikelihood += Math.Log(dist.Density(value));
			}
			return -logLikelihood;
		}
	}

	public class ExcelService
	{
		public static List<MeasurementData> ReadFromExcel(string filePath)
		{
			var data = new List<MeasurementData>();

			ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

			using (var package = new ExcelPackage(new FileInfo(filePath)))
			{
				var worksheet = package.Workbook.Worksheets[0];

				var row = 2; // Начинаем со второй строки (первая — заголовок)
				while (worksheet.Cells[row, 1].Value != null)
				{
					// Считываем Timestamp в формате DateTime
					var timestampRaw = worksheet.Cells[row, 1].Value;
					DateTime timestamp;

					// Пробуем преобразовать значение
					if (DateTime.TryParse(timestampRaw.ToString(), out timestamp))
					{
						// Приводим к краткому формату (для отображения, если потребуется)
						var formattedTimestamp = timestamp.ToString("dd.MM.yyyy HH:mm");

						// Выводим только в формате DateTime в объект MeasurementData
						data.Add(new MeasurementData
						{
							Timestamp = timestamp,
							Value = double.Parse(worksheet.Cells[row, 2].Text)
						});
					}
					else
					{
						throw new FormatException($"Неверный формат даты: {timestampRaw} в строке {row}");
					}

					row++;
				}
			}

			return data;
		}

		public static void SaveToExcel(string filePath, List<ClusterResult> results)
		{
			ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

			using (var package = new ExcelPackage(new FileInfo(filePath)))
			{
				var worksheet = package.Workbook.Worksheets.Add("Results");

				// Заголовки
				worksheet.Cells[1, 1].Value = "Cluster ID";
				worksheet.Cells[1, 2].Value = "Weight";
				worksheet.Cells[1, 3].Value = "Distribution";
				worksheet.Cells[1, 4].Value = "Parameters";

				// Данные
				var row = 2;
				foreach (var result in results)
				{
					worksheet.Cells[row, 1].Value = result.ClusterId;
					worksheet.Cells[row, 2].Value = result.Weight;
					worksheet.Cells[row, 3].Value = result.Distribution;
					worksheet.Cells[row, 4].Value = string.Join(", ", result.Parameters.Select(p => $"{p.Key}: {p.Value}"));
					row++;
				}

				package.SaveAs(new FileInfo(filePath));
			}
		}
	}

	class Program
	{
		static void Main()
		{
			Console.Write("Введите путь к файлу Excel с данными: ");
			var inputFilePath = Console.ReadLine();

			Console.Write("Введите путь для сохранения результатов (Excel): ");
			var outputFilePath = Console.ReadLine();

			Console.Write("Введите количество кластеров: ");
			int clusterCount = int.Parse(Console.ReadLine() ?? "3");

			// Считываем данные из Excel
			var data = ExcelService.ReadFromExcel(inputFilePath).Select(d => d.Value).ToList();
			if (!data.Any())
			{
				Console.WriteLine("Нет данных для анализа.");
				return;
			}

			// Анализ с использованием кластеризации
			var results = DistributionAnalyzer.AnalyzeWithKMeans(data, clusterCount);

			// Сохранение результатов в Excel
			ExcelService.SaveToExcel(outputFilePath, results);

			Console.WriteLine("Анализ завершён. Результаты сохранены в Excel.");
		}
	}
}

