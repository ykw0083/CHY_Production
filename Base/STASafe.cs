using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace FT_ADDON
{
    abstract class STASafe
    {
        Action action;

        public void Execute()
        {
            if (action.GetInvocationList().Length == 0) throw new Exception("STA Safe action is empty");

            Thread t = new Thread(() =>
            {
                action();
            });
            t.SetApartmentState(ApartmentState.STA);
            t.Start();
            t.Join();
        }

        protected void AddAction(Action action)
        {
            this.action += action;
        }
    }
}
