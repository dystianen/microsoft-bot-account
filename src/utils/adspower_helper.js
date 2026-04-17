const axios = require('axios');
const config = require('../config');

class AdsPowerHelper {
  constructor() {
    this.baseUrl = (config.adsPower.baseUrl || '').replace(/\/+$/, '');
  }

  randomInt(min, max) {
    return Math.floor(Math.random() * (max - min + 1)) + min;
  }

  async checkConnection() {
    try {
      await axios.get(`${this.baseUrl}/api/v1/status`, {
        timeout: config.hardTimeout,
        headers: {
          Authorization: `Bearer ${config.adsPower.apiKey}`,
          'api-key': config.adsPower.apiKey,
        },
      });
      return true;
    } catch (error) {
      // If the server responds with ANY status code (404, 401, etc.), the connection exists.
      if (error.response) {
        return true;
      }
      console.error(`[AdsPower Check] Connection failed to ${this.baseUrl}:`, error.message);
      return false;
    }
  }

  randomUserAgent() {
    // Gunakan versi Chrome yang lebih recent dan realistis
    const chromeVersions = ['124', '125', '126', '127', '128', '145', '146'];
    const version = chromeVersions[Math.floor(Math.random() * chromeVersions.length)];
    const minor = this.randomInt(0, 9999);
    const patch = this.randomInt(0, 999);
    return `Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/${version}.0.${minor}.${patch} Safari/537.36`;
  }

  async createProfile(profileName, proxyOverride = null) {
    try {
      const proxy = {
        proxy_type: 'socks5',
        proxy_host: config.proxy.host,
        proxy_port: config.proxy.port,
        proxy_user: proxyOverride?.username || config.proxy.username,
        proxy_password: proxyOverride?.password || config.proxy.password,
        proxy_soft: 'other',
      };

      const response = await axios.post(
        `${this.baseUrl}/api/v1/user/create`,
        {
          name: profileName,
          group_id: config.adsPower.groupId,
          domain_name: '',
          os_type: 'windows',
          repeat_config: ['0'],
          user_proxy_config: proxy,
          fingerprint_config: {
            automatic_timezone: '1',
            language: ['en-US', 'en'],
            flash: 'block',
            fonts: ['all'],
            webrtc: 'disabled',
            canvas: '0',
            webgl: '0',
            webgl_image: '0',
            hardware_concurrency: 'default',
            ua: this.randomUserAgent(),
          },
        },
        {
          headers: {
            Authorization: `Bearer ${config.adsPower.apiKey}`,
            'api-key': config.adsPower.apiKey,
          },
        }
      );

      if (response.data.code !== 0) {
        throw new Error(`AdsPower Profile Creation Failed: ${response.data.msg}`);
      }

      return response.data.data.id;
    } catch (error) {
      if (error.code === 'ECONNREFUSED' || error.code === 'ENOTFOUND') {
        throw new Error(
          'Could not connect to AdsPower Local API. Please ensure AdsPower is open and running on your VPS.'
        );
      }
      if (error.response) {
        console.error('AdsPower Error Response:', JSON.stringify(error.response.data, null, 2));
      }
      console.error('Error creating AdsPower profile:', error.message);
      throw error;
    }
  }

  async startBrowser(profileId, headlessOverride = null) {
    try {
      const isHeadless = headlessOverride !== null ? headlessOverride : config.headless;

      // Menggunakan standar API V2 sesuai dokumentasi USER
      const launchArgs = [
        '--no-sandbox',
        '--disable-gpu',
        '--disable-notifications',
        `--display=${process.env.DISPLAY || ':20'}`,
      ];

      const response = await axios.post(
        `${this.baseUrl}/api/v2/browser-profile/start`,
        {
          profile_id: profileId,
          headless: isHeadless ? '1' : '0',
          launch_args: launchArgs,
          last_opened_tabs: '1',
          proxy_detection: '1',
          password_filling: '0',
          password_saving: '0',
          cdp_mask: '1',
          delete_cache: '0',
          device_scale: '1',
        },
        {
          headers: {
            Authorization: `Bearer ${config.adsPower.apiKey}`,
            'api-key': config.adsPower.apiKey,
            'Content-Type': 'application/json',
          },
        }
      );

      if (response.data.code !== 0) {
        throw new Error(`Failed to start AdsPower browser (V2): ${response.data.msg}`);
      }

      const wsData = response.data?.data?.ws || {};
      const wsUrl = wsData.puppeteer || wsData.selenium;

      if (!wsUrl) {
        throw new Error('Could not find WebSocket URL in AdsPower V2 response');
      }

      return {
        wsUrl: wsUrl,
        debugPort: response.data?.data?.debug_port,
      };
    } catch (error) {
      if (error.code === 'ECONNREFUSED' || error.code === 'ENOTFOUND') {
        throw new Error(
          'Could not connect to AdsPower Local API. Please ensure AdsPower is open and running on your VPS.'
        );
      }
      console.error('Error starting browser:', error.message);
      throw error;
    }
  }

  async stopBrowser(profileId) {
    try {
      await axios.get(`${this.baseUrl}/api/v1/browser/stop?user_id=${profileId}`, {
        headers: {
          Authorization: `Bearer ${config.adsPower.apiKey}`,
          'api-key': config.adsPower.apiKey,
        },
      });
    } catch (error) {
      console.error('Error stopping browser:', error.message);
    }
  }

  async deleteProfile(profileId) {
    try {
      const response = await axios.post(
        `${this.baseUrl}/api/v1/user/delete`,
        {
          user_ids: [profileId],
        },
        {
          headers: {
            Authorization: `Bearer ${config.adsPower.apiKey}`,
            'api-key': config.adsPower.apiKey,
          },
        }
      );

      if (response.data.code !== 0) {
        throw new Error(`Failed to delete profile: ${response.data.msg}`);
      }
      console.log(`Profile ${profileId} deleted successfully.`);
    } catch (error) {
      console.error(`Error deleting profile ${profileId}:`, error.message);
    }
  }
}

module.exports = new AdsPowerHelper();
