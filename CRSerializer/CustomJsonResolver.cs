using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;
using System;
using System.Reflection;

namespace CRSerializer
{
    // Ref: https://stackoverflow.com/questions/20962316/ignoring-class-members-that-throw-exceptions-when-serializing-to-json
    internal class CustomJsonResolver : DefaultContractResolver
    {
        protected override JsonProperty CreateProperty(MemberInfo member, MemberSerialization memberSerialization)
        {
            JsonProperty property = base.CreateProperty(member, memberSerialization);

            property.ShouldSerialize = instance =>
            {
                try
                {
                    PropertyInfo prop = (PropertyInfo)member;
                    if (prop.CanRead)
                    {
                        prop.GetValue(instance, null);
                        return true;
                    }
                }
                catch (Exception)
                {
                }
                return false;
            };
            return property;
        }
    }
}