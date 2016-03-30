require "omniauth/strategies/oauth2"

module OmniAuth
  module Strategies
    class MicrosoftOffice365 < OmniAuth::Strategies::OAuth2
      option :name, :microsoft_office365

      option :client_options, {
        site:          "https://login.microsoftonline.com",
        authorize_url: "/common/oauth2/v2.0/authorize",
        token_url:     "/common/oauth2/v2.0/token"
      }

      option :authorize_params, {
        scope: "openid email profile https://outlook.office.com/contacts.read",
      }

      uid { raw_info["Id"] }

      info do
        {
          email:        raw_info["EmailAddress"],
          display_name: raw_info["DisplayName"],
          first_name:   raw_info["DisplayName"].split(", ")[1],
          last_name:    raw_info["DisplayName"].split(", ")[0],
          alias:        raw_info["Alias"]
        }
      end

      extra do
        {
          "raw_info" => raw_info
        }
      end

      def raw_info
        @raw_info ||= access_token.get("https://outlook.office.com/api/v2.0/me/").parsed
      end
    end
  end
end
