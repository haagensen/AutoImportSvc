using System.Runtime.InteropServices;
using System.Text;

class FnINI
{
    [DllImport("kernel32")]
    static extern long WritePrivateProfileString(string Section, string Key, string Value, string Path);
    [DllImport("kernel32")]
    static extern int GetPrivateProfileString(string Section, string Key, string Default, StringBuilder retval, int Size, string FilePath);

    public string ReadINI(string ArquivoINI, string Section, string Key)
    {
        var retval = new StringBuilder(255);
        GetPrivateProfileString(Section, Key, "", retval, 255, ArquivoINI);
        return retval.ToString();
    }

    public void WriteINI(string ArquivoINI, string Section, string Key, string Valor)
    {
        WritePrivateProfileString(Section, Key, Valor, ArquivoINI);
    }

}
