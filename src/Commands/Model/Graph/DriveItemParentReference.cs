namespace PnP.PowerShell.Commands.Model.Graph
{
    /// <summary>
    /// Reference to the drive, site and folder containing a drive item
    /// </summary>
    public class DriveItemParentReference
    {
        /// <summary>
        /// Id of the drive containing the item
        /// </summary>
        public string DriveId { get; set; }

        /// <summary>
        /// Type of drive, such as documentLibrary or business
        /// </summary>
        public string DriveType { get; set; }

        /// <summary>
        /// Id of the parent folder
        /// </summary>
        public string Id { get; set; }

        /// <summary>
        /// Path of the parent folder
        /// </summary>
        public string Path { get; set; }

        /// <summary>
        /// Composite id of the site containing the item: hostname,siteCollectionId,webId
        /// </summary>
        public string SiteId { get; set; }
    }
}
