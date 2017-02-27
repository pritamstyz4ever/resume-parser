using System.Collections.Generic;
using System.ComponentModel;

namespace DocumentViewer.Models
{
    /// <summary>
    /// 
    /// </summary>
    public class DocumentViewModel
    {
        /// <summary>
        /// Gets or sets the objective.
        /// </summary>
        /// <value>
        /// The objective.
        /// </value>
        public IList<string> Objective { get; set; }
        /// <summary>
        /// Gets or sets the experience summary.
        /// </summary>
        /// <value>
        /// The experience summary.
        /// </value>
        public IList<string> ExperienceSummary { get; set; }
        /// <summary>
        /// Gets or sets the certification.
        /// </summary>
        /// <value>
        /// The certification.
        /// </value>
        public IList<string> Certification { get; set; }
        /// <summary>
        /// Gets or sets the education.
        /// </summary>
        /// <value>
        /// The education.
        /// </value>
        public IList<string> Education { get; set; }
        /// <summary>
        /// Gets or sets the technical skills.
        /// </summary>
        /// <value>
        /// The technical skills.
        /// </value>
        public IList<string> Skill { get; set; }
        /// <summary>
        /// Gets or sets the projects.
        /// </summary>
        /// <value>
        /// The projects.
        /// </value>
        public IList<string> ProjectExperience { get; set; }
        /// <summary>
        /// Gets or sets the achievement.
        /// </summary>
        /// <value>
        /// The achievement.
        /// </value>
        public IList<string> Achievement { get; set; }
        /// <summary>
        /// Gets or sets the links.
        /// </summary>
        /// <value>
        /// The links.
        /// </value>
        public IList<string> Handle { get; set; }

        /// <summary>
        /// Gets or sets the declaration.
        /// </summary>
        /// <value>
        /// The declaration.
        /// </value>
        public IList<string> Declaration { get; set; }

        /// <summary>
        /// Gets or sets the personal.
        /// </summary>
        /// <value>
        /// The personal.
        /// </value>
        public IList<string> Personal { get; set; }
    }
}