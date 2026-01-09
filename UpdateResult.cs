using System;
using System.Collections.Generic;

namespace CadFilesUpdater
{
    public class UpdateResult
    {
        public int TotalFiles { get; set; }
        public int ProcessedFiles { get; set; }
        public int SuccessfulFiles { get; set; }
        public int FailedFiles { get; set; }
        public List<FileError> Errors { get; set; }

        public UpdateResult()
        {
            Errors = new List<FileError>();
        }
    }

    public class FileError
    {
        public string FilePath { get; set; }
        public string ErrorMessage { get; set; }

        public FileError(string filePath, string errorMessage)
        {
            FilePath = filePath;
            ErrorMessage = errorMessage;
        }
    }
}
