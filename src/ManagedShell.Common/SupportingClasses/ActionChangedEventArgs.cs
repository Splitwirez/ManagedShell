using System;
using ManagedShell.Common.Enums;

namespace ManagedShell.Common.SupportingClasses
{
    public class ActionChangedEventArgs : EventArgs
    {
        public readonly Action Action;
        public ActionChangedEventArgs(Action action)
        {
            Action = action;
        }
    }
}
