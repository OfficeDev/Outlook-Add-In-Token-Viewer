// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE in the project root for license information.

namespace TokenValidationService.Models
{
    /// <summary>
    /// Represents the Exchange authentication metadata retrieved from the server
    /// </summary>
    public class ExchangeAuthMetadata
    {
        /// <summary>
        /// The ID of the metadata
        /// </summary>
        public string Id { get; set; }

        /// <summary>
        /// Metadata version
        /// </summary>
        public string Version { get; set; }

        /// <summary>
        /// Metadata name
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Metadata realm
        /// </summary>
        public string Realm { get; set; }

        /// <summary>
        /// Metadata service name
        /// </summary>
        public string ServiceName { get; set; }

        /// <summary>
        /// Metadata issuer
        /// </summary>
        public string Issuer { get; set; }

        /// <summary>
        /// Metadata allowed audiences
        /// </summary>
        public string[] AllowedAudiences { get; set; }

        /// <summary>
        /// Available signing keys
        /// </summary>
        public ExchangeKey[] Keys { get; set; }

    }
}