// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE in the project root for license information.
using System.Configuration;
using System.Threading.Tasks;
using System.Web.Http;
using TokenValidationService.Models;

namespace TokenValidationService.Controllers
{
    /// <summary>
    /// Validates an Office add-in SSO token
    /// </summary>
    public class ValidateSsoTokenController : ApiController
    {
        // POST api/ValidateSsoToken
        /// <summary>
        /// Validates an Exchange user identity token
        /// </summary>
        /// <param name="token">The SSO token as obtained via Office.context.auth.getAccessTokenAsync</param>
        public async Task<SsoTokenValidationResult> Post([FromBody]string token)
        {
            string expectedAudience = ConfigurationManager.AppSettings["ida:AppId"];
            AddInSsoToken ssoToken = new AddInSsoToken(token);

            var result = await ssoToken.Validate(expectedAudience);

            return result;
        }
    }
}