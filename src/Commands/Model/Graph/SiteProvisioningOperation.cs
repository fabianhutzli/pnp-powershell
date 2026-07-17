using System;
using System.Text.Json.Serialization;

namespace PnP.PowerShell.Commands.Model.Graph
{
    /// <summary>
    /// Model for a richLongRunningOperation returned by the Microsoft Graph site provisioning endpoints
    /// </summary>
    public class SiteProvisioningOperation
    {
        /// <summary>
        /// Unique identifier of the operation
        /// </summary>
        [JsonPropertyName("id")]
        public string Id { get; set; }

        /// <summary>
        /// Time when this operation was created
        /// </summary>
        [JsonPropertyName("createdDateTime")]
        public DateTime? CreatedDateTime { get; set; }

        /// <summary>
        /// Time when the last action was performed on this operation
        /// </summary>
        [JsonPropertyName("lastActionDateTime")]
        public DateTime? LastActionDateTime { get; set; }

        /// <summary>
        /// Status of the operation. Possible values are: notStarted, running, succeeded, failed, skipped, unknownFutureValue
        /// </summary>
        [JsonPropertyName("status")]
        public string Status { get; set; }

        /// <summary>
        /// Detail about the status value
        /// </summary>
        [JsonPropertyName("statusDetail")]
        public string StatusDetail { get; set; }

        /// <summary>
        /// Unique identifier of the site that was created, once the operation has succeeded
        /// </summary>
        [JsonPropertyName("resourceId")]
        public string ResourceId { get; set; }

        /// <summary>
        /// Canonical Microsoft Graph URL of the site that was created, once the operation has succeeded
        /// </summary>
        [JsonPropertyName("resourceLocation")]
        public string ResourceLocation { get; set; }

        /// <summary>
        /// A value between 0 and 100 that indicates the progress of the operation
        /// </summary>
        [JsonPropertyName("percentageComplete")]
        public int? PercentageComplete { get; set; }

        /// <summary>
        /// Error due to which the operation failed
        /// </summary>
        [JsonPropertyName("error")]
        public SiteProvisioningOperationError Error { get; set; }
    }

    /// <summary>
    /// Model for the error returned as part of a failed SiteProvisioningOperation
    /// </summary>
    public class SiteProvisioningOperationError
    {
        [JsonPropertyName("code")]
        public string Code { get; set; }

        [JsonPropertyName("message")]
        public string Message { get; set; }
    }
}
