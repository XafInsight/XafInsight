namespace xafplugin.Interfaces
{
    public interface IEnvironmentService
    {
        string DatabasePath { get; set; }
        string FileHash { get; set; }
    }
}
