using System;

namespace ACUtils.AXRepository.Exceptions
{
    internal class AXFieldNotFoundException: Exception
    {
        public AXFieldNotFoundException(string message): base(message)
        {
        }
    }
}
