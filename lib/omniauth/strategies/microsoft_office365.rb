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
          image:        avatar_file,
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

      private

      def avatar_file
        photo = access_token.get("https://outlook.office.com/api/v2.0/me/photo/$value")
        ext   = photo.content_type.sub("image/", "") # "image/jpeg" => "jpeg"

        Tempfile.new(["avatar", ".#{ext}"]).tap do |file|
          file.binmode
          file.write(photo.body)
          file.rewind
        end

      rescue ::OAuth2::Error => e
        if e.response.status == 404 # User has no avatar...
          return nil
        else
          raise
        end
      end

    end
  end
end
