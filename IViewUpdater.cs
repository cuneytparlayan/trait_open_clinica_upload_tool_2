using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OCDataImporter
{
    /// <summary>
    /// Defines the methods a class must implement to update external view components from 
    /// the model or controller.
    /// </summary>
    interface IViewUpdater
    {
        /// <summary>
        /// Must indicate the step size to a progress bar
        /// </summary>
        /// <param name="step"></param>
        void updateProgressbarStep(int step);

        /// <summary>
        /// Must perform a progressbar step        
        /// </summary>
        void performProgressbarStep();

        /// <summary>
        /// Must append a messasge to a text output box. E.g. with information on the conversion
        /// </summary>
        void appendText(String aMessage);

        /// <summary>
        /// Must reset the messasge in a text output box 
        /// </summary>
        void resetText();
    }
}
