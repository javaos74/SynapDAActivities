using System.Activities;
using UiPath.Robot.Activities.Api;

namespace Charles.Synap.Activities.Helpers
{
    public static class ActivityContextExtensions
    {
        public static IExecutorRuntime GetExecutorRuntime(this ActivityContext context) => context.GetExtension<IExecutorRuntime>();
    }
}
