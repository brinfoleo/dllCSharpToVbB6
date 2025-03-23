
using System.Runtime.InteropServices;

namespace dllBBCobrancaFW35
{
    
    [Guid("4C82DC72-74E3-48E5-9CD7-55F42771A7DF")]
    //[Guid("SEU-GUID-AQUI")] // Gere um GUID único usando o Visual Studio (Ferramentas > Criar GUID)
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    [ComVisible(true)]
    public interface IMinhaInterface
    {
        [DispId(1)]
        //string MinhaFuncao(string texto);
        string MinhaFuncao(string texto);
    }

    //[Guid("OUTRO-GUID-AQUI")] // Gere outro GUID único
    
    [Guid("77C3DF46-2567-4247-9312-C4A3E7E0A086")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComVisible(true)]
    public class BBCob : IMinhaInterface
    {

        public string MinhaFuncao(string texto)
        {
            return "VB6 disse: " + texto;
        }
    }
}
