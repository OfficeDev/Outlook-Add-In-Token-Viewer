// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE in the project root for license information.
using System.Web.Http;
using TokenValidationService.Models;

namespace TokenValidationService.Controllers
{
    /// <summary>
    /// Validates an Exchange user identity token
    /// </summary>
    public class ValidateExchangeTokenController : ApiController
    {
        
        // POST api/ValidateExchangeToken
        /// <summary>
        /// Validates an Exchange user identity token
        /// </summary>
        /// <param name="token">The Exchange user identity token as obtained via Office.context.mailbox.getUserIdentityTokenAsync</param>
        public IdTokenValidationResult Post([FromBody]string token)
        {
            ExchangeIdToken idToken = new ExchangeIdToken(token);

            var result = idToken.Validate("https://localhost:44359/add-in/TaskPane/TaskPane.html");

            return result;
        }
    }
}
