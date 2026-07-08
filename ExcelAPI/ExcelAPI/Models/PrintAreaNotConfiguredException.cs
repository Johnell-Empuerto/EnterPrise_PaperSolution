namespace ExcelAPI.Models
{
    /// <summary>
    /// Exception thrown when the uploaded Excel worksheet does not have a print area configured.
    /// This is a user error that should result in a 400 Bad Request response.
    /// </summary>
    public class PrintAreaNotConfiguredException : Exception
    {
        public PrintAreaNotConfiguredException(string message) : base(message)
        {
        }

        public PrintAreaNotConfiguredException(string message, Exception innerException)
            : base(message, innerException)
        {
        }
    }
}
