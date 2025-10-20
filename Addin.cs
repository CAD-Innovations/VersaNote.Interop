using System.Linq;
using System.Reflection;

namespace VersaNote.Interop
{
    public class Addin
    {
        private object VersaNoteObject;
        private MethodInfo[] methods;

        public Addin(object versaNoteObject)
        {
            VersaNoteObject = versaNoteObject;
            methods = VersaNoteObject.GetType().GetMethods(BindingFlags.Public | BindingFlags.Instance);
        }

        public void ShowToastNotification(string title, string message)
        {
            MethodInfo method = methods.FirstOrDefault(x => x.Name == nameof(ShowToastNotification));
            if (method != null)
            {
                object[] parameters = new object[] { title, message };
                method.Invoke(VersaNoteObject, parameters);
            }

        }
    }
}
