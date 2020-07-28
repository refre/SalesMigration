using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Input;
using System.Diagnostics;

namespace Ista.Migration.Commands
{
    /// <summary>
    /// This class is a typical relay command coming from a microsost example.
    /// http://msdn.microsoft.com/en-us/magazine/dd419663.aspx
    /// </summary>
    class RelayCommand:ICommand
    {
        /// <summary>
        /// private execute field.
        /// </summary>
        readonly Action<object> _execute;
        /// <summary>
        /// private can execute field.
        /// </summary
        readonly Predicate<Object> _canExecute;
        /// <summary>
        /// Initialize a new relaycommand.
        /// </summary>
        /// <param name="execute"></param>
        public RelayCommand(Action<object> execute) : this(execute,null) { }
        /// <summary>
        /// Initialize a new relaycommand.
        /// </summary>
        /// <param name="execute"></param>
        /// <param name="canExecute"></param>
        public RelayCommand(Action<object> execute, Predicate<object> canExecute)
        {
            _execute = execute;
            _canExecute = canExecute;
        }


        /// <summary>
        /// Check whether an action could be executed.
        /// </summary>
        /// <param name="parameter"></param>
        /// <returns></returns>
        [DebuggerStepThrough]
        public bool CanExecute(object parameter)
        {
            return _canExecute == null ? true : _canExecute(parameter);
        }

        public event EventHandler CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }
        /// <summary>
        /// Execute an action could.
        /// </summary>
        public void Execute(object parameter)
        {
            _execute(parameter);
        }
    }
}
