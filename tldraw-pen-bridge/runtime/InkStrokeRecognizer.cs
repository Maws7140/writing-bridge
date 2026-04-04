using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Web.Script.Serialization;
using Microsoft.Ink;

namespace TldrawPenBridgeRuntime
{
    public sealed class RecognizerRequest
    {
        public string Mode { get; set; }
        public List<StrokeInput> Strokes { get; set; }
        public int AlternateCount { get; set; }
    }

    public sealed class StrokeInput
    {
        public List<PointInput> Points { get; set; }
    }

    public sealed class PointInput
    {
        public double X { get; set; }
        public double Y { get; set; }
        public double Z { get; set; }
    }

    public sealed class RecognizerResponse
    {
        public bool Ok { get; set; }
        public string Text { get; set; }
        public string Status { get; set; }
        public string Confidence { get; set; }
        public List<string> Alternates { get; set; }
        public int StrokeCount { get; set; }
        public string Error { get; set; }
    }

    internal static class Program
    {
        private const double TargetMaxDimension = 1400.0;
        private const double Margin = 60.0;

        private static int Main(string[] args)
        {
            var serializer = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };

            try
            {
                if (args.Length != 1 || string.IsNullOrWhiteSpace(args[0]))
                {
                    Write(serializer, new RecognizerResponse
                    {
                        Ok = false,
                        Error = "Expected a single JSON input path."
                    });
                    return 1;
                }

                var inputPath = args[0];
                if (!File.Exists(inputPath))
                {
                    Write(serializer, new RecognizerResponse
                    {
                        Ok = false,
                        Error = "Input file does not exist."
                    });
                    return 1;
                }

                var request = serializer.Deserialize<RecognizerRequest>(File.ReadAllText(inputPath));
                var response = Recognize(request);
                Write(serializer, response);
                return response.Ok ? 0 : 1;
            }
            catch (Exception ex)
            {
                Write(serializer, new RecognizerResponse
                {
                    Ok = false,
                    Error = ex.Message
                });
                return 1;
            }
        }

        private static RecognizerResponse Recognize(RecognizerRequest request)
        {
            if (request == null || request.Strokes == null || request.Strokes.Count == 0)
            {
                return new RecognizerResponse
                {
                    Ok = false,
                    Error = "No strokes were provided."
                };
            }

            var strokePointSets = request.Strokes
                .Select(stroke =>
                {
                    var points = stroke != null && stroke.Points != null ? stroke.Points : new List<PointInput>();
                    return points.Select(point => new PointF((float)point.X, (float)point.Y)).ToList();
                })
                .Where(points => points.Count > 0)
                .ToList();

            if (strokePointSets.Count == 0)
            {
                return new RecognizerResponse
                {
                    Ok = false,
                    Error = "No stroke points were provided."
                };
            }

            var normalized = Normalize(strokePointSets);
            using (var ink = new Ink())
            using (var context = new RecognizerContext())
            {
                foreach (var strokePoints in normalized)
                {
                    ink.CreateStroke(strokePoints);
                }

                context.Strokes = ink.Strokes;
                ApplyFactoid(context, request.Mode);

                RecognitionStatus status;
                var result = context.Recognize(out status);
                if (result == null)
                {
                    return new RecognizerResponse
                    {
                        Ok = false,
                        Status = status.ToString(),
                        Error = "Recognizer returned no result.",
                        StrokeCount = normalized.Count
                    };
                }

                var alternates = new List<string>();
                try
                {
                    var maxAlternates = request.AlternateCount > 0 ? request.AlternateCount : 5;
                    var allAlternates = result.GetAlternatesFromSelection();
                    for (var i = 0; i < allAlternates.Count && alternates.Count < maxAlternates; i++)
                    {
                        var alternate = allAlternates[i];
                        if (alternate == null || string.IsNullOrWhiteSpace(alternate.ToString()))
                        {
                            continue;
                        }

                        var candidate = alternate.ToString().Trim();
                        if (!alternates.Contains(candidate, StringComparer.OrdinalIgnoreCase))
                        {
                            alternates.Add(candidate);
                        }
                    }
                }
                catch
                {
                }

                var text = string.IsNullOrWhiteSpace(result.TopString) ? null : result.TopString.Trim();
                if (string.IsNullOrWhiteSpace(text) && alternates.Count > 0)
                {
                    text = alternates[0];
                }

                return new RecognizerResponse
                {
                    Ok = !string.IsNullOrWhiteSpace(text),
                    Text = text,
                    Status = status.ToString(),
                    Confidence = result.TopAlternate != null ? result.TopAlternate.Confidence.ToString() : null,
                    Alternates = alternates,
                    StrokeCount = normalized.Count,
                    Error = string.IsNullOrWhiteSpace(text) ? "Recognizer produced no text." : null
                };
            }
        }

        private static void ApplyFactoid(RecognizerContext context, string mode)
        {
            if (string.IsNullOrWhiteSpace(mode))
            {
                return;
            }

            switch (mode.Trim().ToLowerInvariant())
            {
                case "number":
                    context.Factoid = Factoid.Number;
                    break;
                case "digit":
                    context.Factoid = Factoid.Digit;
                    break;
                default:
                    break;
            }
        }

        private static List<Point[]> Normalize(List<List<PointF>> strokes)
        {
            var minX = strokes.SelectMany(points => points).Min(point => point.X);
            var minY = strokes.SelectMany(points => points).Min(point => point.Y);
            var maxX = strokes.SelectMany(points => points).Max(point => point.X);
            var maxY = strokes.SelectMany(points => points).Max(point => point.Y);
            var width = Math.Max(1, maxX - minX);
            var height = Math.Max(1, maxY - minY);
            var scale = TargetMaxDimension / Math.Max(width, height);
            scale = Math.Max(0.75, Math.Min(scale, 8.0));

            var output = new List<Point[]>();
            foreach (var stroke in strokes)
            {
                var normalized = new List<Point>();
                Point? lastPoint = null;

                foreach (var point in stroke)
                {
                    var current = new Point(
                        (int)Math.Round((point.X - minX) * scale + Margin),
                        (int)Math.Round((point.Y - minY) * scale + Margin)
                    );

                    if (lastPoint.HasValue && lastPoint.Value == current)
                    {
                        continue;
                    }

                    normalized.Add(current);
                    lastPoint = current;
                }

                if (normalized.Count == 1)
                {
                    var only = normalized[0];
                    normalized.Add(new Point(only.X + 1, only.Y + 1));
                }

                if (normalized.Count >= 2)
                {
                    output.Add(normalized.ToArray());
                }
            }

            return output;
        }

        private static void Write(JavaScriptSerializer serializer, RecognizerResponse response)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            Console.WriteLine(serializer.Serialize(response));
        }
    }
}
