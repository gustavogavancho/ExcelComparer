using System.ComponentModel;

namespace ExcelComparer.WPFUI.Core;

public class ViewModelBase : INotifyPropertyChanged
{
    public virtual void Dispose() { }

    public event PropertyChangedEventHandler PropertyChanged;

    protected void OnPropertyChanged(string propertyName)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}
