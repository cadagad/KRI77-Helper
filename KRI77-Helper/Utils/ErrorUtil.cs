using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KRI77_Helper.Utils
{
    internal class ErrorUtil
    {

        // Map allowed file extensions per process.
        // ✅ Edit this as your processes evolve.
        private static readonly Dictionary<string, string[]> AllowedExtensionsByProcess =
            new(StringComparer.OrdinalIgnoreCase)
            {
                ["Printer"] = new[] { ".csv" },             //Printers
                ["Servers"] = new[] { ".csv"},              //Tanium Servers
                ["NetworkNA"] = new[] { ".xlsx" },          //Network
                ["NetworkAsia"] = new[] { ".csv" },         //Network
                ["Intune Report"] = new[] { ".csv" },       //Intune Report
                ["End User Devices"] = new[] { ".csv" },    //Tanium EUD
                ["Terminal"] = new[] { ".csv" },            //Terminals

                // Fallback if processName not mapped
                ["Default"] = new[] { ".csv", ".xlsx" }
            };

        /// <summary>
        /// MAIN ERROR FUNCTION (spec #5)
        /// Validates filename and file type. Returns "Success" or "Error: ..." (combined messages).
        /// </summary>
        //public static string HandleProcessInputError(string processName, string inputFile, string expectedFilename)
        public static string HandleProcessInputError(string processName, string inputFile)
        {
            var errors = new List<string>();

            // Derive the actual filename from the path, then call the filename validator
            //var actualFileName = Path.GetFileName(inputFile);

            //if (!ValidateIncorrectFilename(actualFileName, expectedFilename, out var nameErr))
            //    errors.Add(nameErr);

            if (!ValidateIncorrectFileType(processName, inputFile, out var typeErr))
                errors.Add(typeErr);

            return errors.Count == 0
                ? "Success"
                : "Error: " + string.Join("; ", errors);
        }



        /// <summary>
        /// CHECK #1: Incorrect filename
        /// Parameters per your spec:
        /// - sourceFile: the actual filename (e.g., "printers.csv")
        /// - expectedFilename: the expected filename (e.g., "printers.csv")
        /// Returns false and sets error if they don't match.
        /// </summary>
        public static bool ValidateIncorrectFilename(string sourceFile, string expectedFilename, out string error)
        {
            error = string.Empty;

            if (string.IsNullOrWhiteSpace(sourceFile))
            {
                error = "Actual filename (sourceFile) is empty.";
                return false;
            }

            if (string.IsNullOrWhiteSpace(expectedFilename))
            {
                error = "Expected filename is empty.";
                return false;
            }

            // Be tolerant if a path is passed accidentally; compare only the names.
            var actual = Path.GetFileName(sourceFile);
            var expected = Path.GetFileName(expectedFilename);

            // Case-insensitive compare (typical for Windows filesystems)
            if (!string.Equals(actual, expected, StringComparison.OrdinalIgnoreCase))
            {
                error = $"Incorrect filename. Expected '{expected}', got '{actual}'.";
                return false;
            }

            return true;
        }


        /// <summary>
        /// CHECK #2: Incorrect file type
        /// - Extracts extension from inputFilePath
        /// - Compares against allowed extensions for the given process
        /// - Uses "Default" set if process is not mapped
        /// </summary>
        public static bool ValidateIncorrectFileType(string processName, string inputFilePath, out string error)
        {
            error = string.Empty;

            if (string.IsNullOrWhiteSpace(inputFilePath))
            {
                error = "File path is empty.";
                return false;
            }

            var ext = Path.GetExtension(inputFilePath);
            if (string.IsNullOrWhiteSpace(ext))
            {
                error = "File has no extension.";
                return false;
            }

            // Get allowed extensions for this process, or fallback to Default
            if (!AllowedExtensionsByProcess.TryGetValue(processName ?? string.Empty, out var allowed))
            {
                allowed = AllowedExtensionsByProcess["Default"];
            }

            if (!allowed.Contains(ext, StringComparer.OrdinalIgnoreCase))
            {
                error = $"Incorrect file type for process '{processName}'. " +
                        $"Allowed: {string.Join(", ", allowed)}. Provided: '{ext}'.";
                return false;
            }

            return true;
        }


    }
}
