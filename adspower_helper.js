const axios = require("axios");
const config = require("./config");

class AdsPowerHelper {
  constructor() {
    this.baseUrl = (config.adsPower.baseUrl || "").replace(/\/+$/, "");
  }

  randomInt(min, max) {
    return Math.floor(Math.random() * (max - min + 1)) + min;
  }

  randomUserAgent() {
    // Gunakan versi Chrome yang lebih recent dan realistis
    const chromeVersions = ["124", "125", "126", "127", "128"];
    const version =
      chromeVersions[Math.floor(Math.random() * chromeVersions.length)];
    const minor = this.randomInt(0, 9999);
    const patch = this.randomInt(0, 999);
    return `Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/${version}.0.${minor}.${patch} Safari/537.36`;
  }

  async createProfile(profileName) {
    try {
      const response = await axios.post(
        `${this.baseUrl}/api/v1/user/create`,
        {
          name: profileName,
          group_id: config.adsPower.groupId,
          domain_name: "",
          os_type: "windows",
          repeat_config: ["0"],
          user_proxy_config: {
            proxy_type: "socks5",
            proxy_host: config.proxy.host,
            proxy_port: config.proxy.port,
            proxy_user: config.proxy.username,
            proxy_password: config.proxy.password,
            proxy_soft: "other",
          },
          fingerprint_config: {
            automatic_timezone: "1",
            // Sesuaikan bahasa dengan lokasi proxy (Indonesia)
            // Ini penting — mismatch bahasa vs IP lokasi bisa trigger CAPTCHA
            language: ["id-ID", "id", "en-US", "en", "ja-JP", "ja", "ko-KR", "ko", "zh-CN", "zh", "zh-TW", "zh-HK", "zh-MO", "zh-SG", "zh-Hans", "zh-Hant", "fr-FR", "fr", "de-DE", "de", "it-IT", "it", "es-ES", "es", "pt-PT", "pt", "ru-RU", "ru", "ar-SA", "ar", "hi-IN", "hi", "bn-BD", "bn", "ta-IN", "ta", "te-IN", "te", "ur-PK", "ur", "fa-IR", "fa", "th-TH", "th", "vi-VN", "vi", "ko-KR", "ko", "zh-CN", "zh", "zh-TW", "zh-HK", "zh-MO", "zh-SG", "zh-Hans", "zh-Hant"],
            flash: "block",
            fonts: ["all"],
            webrtc: "disabled",
            canvas: "1",       // Sedikit noise lebih natural dari "0"
            webgl: "1",        // Sama
            webgl_image: "1",
            hardware_concurrency: "default",
            ua: this.randomUserAgent(),
          },
        },
        {
          headers: {
            Authorization: `Bearer ${config.adsPower.apiKey}`,
            "api-key": config.adsPower.apiKey,
          },
        },
      );

      if (response.data.code !== 0) {
        throw new Error(
          `AdsPower Profile Creation Failed: ${response.data.msg}`,
        );
      }

      return response.data.data.id;
    } catch (error) {
      if (error.response) {
        console.error(
          "AdsPower Error Response:",
          JSON.stringify(error.response.data, null, 2),
        );
      }
      console.error("Error creating AdsPower profile:", error.message);
      throw error;
    }
  }

  async startBrowser(profileId) {
    try {
      const headlessParam = config.headless ? "&open_tabs=1&headless=1" : "";
      const response = await axios.get(
        `${this.baseUrl}/api/v1/browser/start?user_id=${profileId}${headlessParam}`,
        {
          headers: {
            Authorization: `Bearer ${config.adsPower.apiKey}`,
            "api-key": config.adsPower.apiKey,
          },
        },
      );

      if (response.data.code !== 0) {
        throw new Error(
          `Failed to start AdsPower browser: ${response.data.msg}`,
        );
      }

      const wsData = response.data?.data?.ws || {};
      const wsUrl = wsData.puppeteer || wsData.selenium;

      if (!wsUrl) {
        throw new Error("Could not find WebSocket URL in AdsPower response");
      }

      return {
        wsUrl: wsUrl,
        debugPort: response.data?.data?.debug_port,
      };
    } catch (error) {
      console.error("Error starting browser:", error.message);
      throw error;
    }
  }

  async stopBrowser(profileId) {
    try {
      await axios.get(
        `${this.baseUrl}/api/v1/browser/stop?user_id=${profileId}`,
        {
          headers: {
            Authorization: `Bearer ${config.adsPower.apiKey}`,
            "api-key": config.adsPower.apiKey,
          },
        },
      );
    } catch (error) {
      console.error("Error stopping browser:", error.message);
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
            "api-key": config.adsPower.apiKey,
          },
        },
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