using Papercut.Core.Infrastructure.Plugins;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Papercut.Module.Exchange
{
    using Autofac;
    using Autofac.Core;
    using Core.Domain.Rules;
    using System.Reflection;
    using Module = Autofac.Module;

    public class ExchangePluginModule : Module, IPluginModule
    {
        public IModule Module => this;
        public string Name => "Exchange";
        public string Version => "1.0.0";
        public string Description => "Provides a forwarding mails to Exchange.";

        protected override void Load(ContainerBuilder builder)
        {
            // rules and rule dispatchers
            builder.RegisterAssemblyTypes(Assembly.GetExecutingAssembly())
                .AssignableTo<IRule>()
                .As<IRule>()
                .InstancePerDependency();

            builder.RegisterAssemblyTypes(Assembly.GetExecutingAssembly())
                .AsClosedTypesOf(typeof(IRuleDispatcher<>))
                .AsImplementedInterfaces()
                .AsSelf()
                .InstancePerDependency();

            base.Load(builder);
        }
    }
}
