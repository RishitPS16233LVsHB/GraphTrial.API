namespace GraphTrial.API.Common
{
    public enum ResponseFlag
    {
        Success = 1,
        Error = 0
    }

    public class ResponseResult
    {
        public object Data { get; set; }
        public ResponseFlag Result { get; set; }
        public string Message { get; set; }
        public string StackTrace { get; set; }
    }
}
