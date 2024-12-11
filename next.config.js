/** @type {import('next').NextConfig} */
const nextConfig = {
    webpack: (config) => {
      config.resolve.alias['@'] = __dirname + '/src';
      return config;
    },
  };
  
  
  module.exports = nextConfig;