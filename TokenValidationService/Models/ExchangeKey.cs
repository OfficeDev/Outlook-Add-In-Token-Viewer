// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE in the project root for license information.
using Newtonsoft.Json;

namespace TokenValidationService.Models
{
    /// <summary>
    /// Represents the "keyinfo" property in an Exchange signing key
    /// </summary>
    public class ExchangeKeyInfo
    {
        /// <summary>
        /// The signing key thumprint
        /// </summary>
        [JsonProperty("x5t")]
        public string Thumbprint { get; set; }
    }

    /// <summary>
    /// Represents the "keyvalue" property in an Exchange signing key
    /// </summary>
    public class ExchangeKeyValue
    {
        /// <summary>
        /// The type of signing key. Should be "x509Certificate"
        /// </summary>
        public string Type { get; set; }
        /// <summary>
        /// The base64-encoded signing key certificate
        /// </summary>
        public string Value { get; set; }
    }

    /// <summary>
    /// Represent an Exchange signing key in the Exchange authentication metadata response
    /// </summary>
    public class ExchangeKey
    {
        /// <summary>
        /// The intended usage for the key
        /// </summary>
        public string Usage { get; set; }

        /// <summary>
        /// Information about the key, including the thumprint
        /// </summary>
        [JsonProperty("keyinfo")]
        public ExchangeKeyInfo KeyInfo { get; set;}

        /// <summary>
        /// The signing key value
        /// </summary>
        [JsonProperty("keyvalue")]
        public ExchangeKeyValue KeyValue { get; set; }

    }
}