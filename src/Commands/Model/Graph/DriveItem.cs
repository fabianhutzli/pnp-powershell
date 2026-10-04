using System;

namespace PnP.PowerShell.Commands.Model.Graph
{
    /// <summary>
    /// A file or folder in a drive, as returned by Microsoft Graph
    /// </summary>
    public class DriveItem
    {
        /// <summary>
        /// The drive item id
        /// </summary>
        public string Id { get; set; }

        /// <summary>
        /// The name of the file or folder
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Size of the item in bytes
        /// </summary>
        public long? Size { get; set; }

        /// <summary>
        /// URL to open the item in the browser
        /// </summary>
        public string WebUrl { get; set; }

        /// <summary>
        /// ETag of the entire item, metadata and content
        /// </summary>
        public string ETag { get; set; }

        /// <summary>
        /// ETag of the content of the item
        /// </summary>
        public string CTag { get; set; }

        /// <summary>
        /// Date and time the item was created
        /// </summary>
        public DateTime? CreatedDateTime { get; set; }

        /// <summary>
        /// Date and time the item was last modified
        /// </summary>
        public DateTime? LastModifiedDateTime { get; set; }

        /// <summary>
        /// Identity of the user, device or application that created the item
        /// </summary>
        public IdentitySet CreatedBy { get; set; }

        /// <summary>
        /// Identity of the user, device or application that last modified the item
        /// </summary>
        public IdentitySet LastModifiedBy { get; set; }

        /// <summary>
        /// The drive, site and folder containing the item
        /// </summary>
        public DriveItemParentReference ParentReference { get; set; }

        /// <summary>
        /// File facet, only present when the item is a file
        /// </summary>
        public DriveItemFile File { get; set; }

        /// <summary>
        /// Folder facet, only present when the item is a folder
        /// </summary>
        public DriveItemFolder Folder { get; set; }
    }
}
