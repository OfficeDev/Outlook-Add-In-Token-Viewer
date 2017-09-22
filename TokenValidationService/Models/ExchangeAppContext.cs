// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE in the project root for license information.
using Newtonsoft.Json;

namespace TokenValidationService.Models
{
    /// <summary>
    /// Representation of the appctx claim in an Exchange user identity token.
    /// </summary>
    public class ExchangeAppContext
    {
        /// <summary>
        /// The Exchange identifier for the user
        /// </summary>
        [JsonProperty("msexchuid")]
        public string ExchangeUid { get; set; }
        
        /// <summary>
        /// The token version
        /// </summary>
        public string Version { get; set; }

        /// <summary>
        /// The URL to download authentication metadata
        /// </summary>
        [JsonProperty("amurl")]
        public string MetadataUrl { get; set; }
    }
}