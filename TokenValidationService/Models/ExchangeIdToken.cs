// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE in the project root for license information.
using Microsoft.IdentityModel.Tokens;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IdentityModel.Tokens.Jwt;
using System.Linq;
using System.Net;
using System.Security.Cryptography.X509Certificates;
using System.Text;

namespace TokenValidationService.Models
{
    /// <summary>
    /// Represents an Exchange user identity token. Provides access to the Exchange-specific claims.
    /// </summary>
    public class ExchangeIdToken : JwtSecurityToken
    {
        /// <summary>
        /// The Exchange-specific claims in the token, stored in the "appctx" claim
        /// </summary>
        public ExchangeAppContext AppContext { get; private set; }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="encodedToken">The serialized JWT Token</param>
        public ExchangeIdToken(string encodedToken) : base (encodedToken)
        {
            // Parse the appctx claim to get Exchange-specific info
            var appctx = Claims.FirstOrDefault(claim => claim.Type == "appctx");
            if (appctx != null)
            {
                AppContext = JsonConvert.DeserializeObject<ExchangeAppContext>(appctx.Value);
            }
        }

        /// <summary>
        /// Validates the token
        /// </summary>
        /// <param name="expectedAudience">The valid audience value to check</param>
        /// <returns></returns>
        public IdTokenValidationResult Validate(string expectedAudience)
        {
            IdTokenValidationResult result = new IdTokenValidationResult();

            // Validate non-standard stuff

            // Is appctx valid?
            if (AppContext == null)
            {
                result.Message = "The app context claim is missing or invalid.";
                return result;
            }

            // Token version
            if (string.Compare(AppContext.Version, "ExIdTok.V1", StringComparison.InvariantCulture) == 0)
            {
                result.VersionResult = "passed";
            }

            // Use System.IdentityModel.Tokens.Jwt library to validate standard parts
            JwtSecurityTokenHandler tokenHandler = new JwtSecurityTokenHandler();
            TokenValidationParameters tvp = new TokenValidationParameters();

            tvp.ValidateIssuer = false;
            tvp.ValidateAudience = true;
            tvp.ValidAudience = expectedAudience;
            tvp.ValidateIssuerSigningKey = true;
            tvp.IssuerSigningKeys = GetSigningKeys();
            tvp.ValidateLifetime = true;

            try
            {
                var claimsPrincipal = tokenHandler.ValidateToken(RawData, tvp, out SecurityToken validatedToken);

                // If no exception, all standard checks passed
                result.LifetimeResult = result.SignatureResult = result.AudienceResult = "passed";
                if (string.Compare(result.VersionResult, "passed", StringComparison.InvariantCulture) == 0)
                {
                    result.IsValid = true;

                    // Compute user ID
                    result.ComputedUserId = GenerateUserId(AppContext.ExchangeUid, AppContext.MetadataUrl);
                }
            }
            catch (SecurityTokenInvalidAudienceException ex)
            {
                result.AudienceResult = "failed";
                result.Message = ex.Message;
            }
            catch (SecurityTokenInvalidLifetimeException ex)
            {
                result.LifetimeResult = "failed";
                result.Message = ex.Message;
            }
            catch (SecurityTokenExpiredException ex)
            {
                result.LifetimeResult = "failed";
                result.Message = ex.Message;
            }
            catch (SecurityTokenInvalidSignatureException ex)
            {
                result.SignatureResult = "failed";
                result.Message = ex.Message;
            }
            catch (SecurityTokenValidationException ex)
            {
                result.Message = ex.Message;
            }

            return result;
        }

        private List<SecurityKey> GetSigningKeys()
        {
            // TODO: Implement a cache of signing keys with the auth metadata URL
            // as an index
            // When requests come in to validate a token, check if you already have cached signing keys
            // for that URL

            // Load tokens
            var webClient = new WebClient();
            var authMetaData = JsonConvert.DeserializeObject<ExchangeAuthMetadata>(webClient.DownloadString(AppContext.MetadataUrl));

            // Build list of signing keys
            List<SecurityKey> signingKeys = new List<SecurityKey>();

            foreach (ExchangeKey key in authMetaData.Keys)
            {
                if (string.Compare(key.KeyInfo.Thumbprint, Header.X5t, StringComparison.InvariantCulture) == 0 &&
                    string.Compare(key.KeyValue.Type, "x509Certificate", StringComparison.InvariantCulture) == 0)
                {
                    signingKeys.Add(new X509SecurityKey(new X509Certificate2(Encoding.UTF8.GetBytes(key.KeyValue.Value))));
                }
            }

            return signingKeys;
        }

        private string GenerateUserId(string exchangeUserId, string authMetadataUrl)
        {
            // Generate a binary user ID just from the concatenation of the Exchange
            // user ID and the auth metadata URL that was used to validate the
            // token signature.

            // In a real world scenario if this user ID was going to be used for any reason
            // other than just to correlate Exchange tokens with a backend service, it would
            // be a good idea to secure this with crypto
            return Convert.ToBase64String(Encoding.UTF8.GetBytes(exchangeUserId + authMetadataUrl));
        }
    }
}