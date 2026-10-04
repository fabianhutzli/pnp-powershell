using System;
using System.Collections;

namespace PnP.PowerShell.Commands.Model.Graph
{
    /// <summary>
    /// A SharePoint list item with its field values, as returned by Microsoft Graph
    /// </summary>
    public class GraphListItem
    {
        /// <summary>
        /// The list item id
        /// </summary>
        public string Id { get; set; }

        /// <summary>
        /// URL to open the item in the browser
        /// </summary>
        public string WebUrl { get; set; }

        /// <summary>
        /// ETag of the item
        /// </summary>
        public string ETag { get; set; }

        /// <summary>
        /// Date and time the item was created
        /// </summary>
        public DateTime? CreatedDateTime { get; set; }

        /// <summary>
        /// Date and time the item was last modified
        /// </summary>
        public DateTime? LastModifiedDateTime { get; set; }

        /// <summary>
        /// The field values of the item, keyed by internal field name
        /// </summary>
        public Hashtable Fields { get; set; }
    }
}
