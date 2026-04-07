'use strict';

/**
 * apv service
 */

const { createCoreService } = require('@strapi/strapi').factories;

module.exports = createCoreService('api::apv.apv');
