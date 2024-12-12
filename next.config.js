/** @type {import('next').NextConfig} */
const nextConfig = {
  output: 'standalone',
  webpack: (config) => {
    config.resolve.alias['@'] = __dirname + '/src';
    return config;
  },
};

module.exports = nextConfig;