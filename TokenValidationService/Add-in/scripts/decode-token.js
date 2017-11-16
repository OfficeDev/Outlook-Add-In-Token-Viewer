// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE in the project root for license information.

"use strict";

/**
 * Takes a JWT encoded token and parses it into
 * an object. This method does not do any validation
 * of the signature.
 * @param {string} encodedToken
 * @returns {any} The parsed token
 */
function decodeToken(encodedToken) {
    if (!encodedToken || encodedToken.length <= 0) {
        return null;
    }

    // Encoded tokens have three parts separated by '.'
    var tokenParts = encodedToken.split(".");
    if (!tokenParts || tokenParts.length !== 3) {
        return null;
    }

    // The payload is the second part
    var encodedPayload = tokenParts[1];
    var decodedPayload = atob(encodedPayload);

    // Parse the decoded payload into an object
    return JSON.parse(decodedPayload);
}