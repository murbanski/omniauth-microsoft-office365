require 'spec_helper'

RSpec.describe OmniAuth::Strategies::MicrosoftOffice365 do
  let(:request) { double('Request', :params => {}, :cookies => {}, :env => {}) }
  let(:options) { { } }

  let(:app) do
    lambda do
      [200, {}, ["Hello."]]
    end
  end

  let(:strategy) do
    OmniAuth::Strategies::MicrosoftOffice365.new(app, 'appid', 'secret', options)
  end

  before do
    OmniAuth.config.test_mode = true
    allow(strategy).to receive(:request).and_return(request)
  end

  after do
    OmniAuth.config.test_mode = false
  end

  describe "#name" do
    it "returns :microsoft_office365" do
      expect(strategy.name).to eq(:microsoft_office365)
    end
  end

  describe "#client_options" do
    context "with defaults" do
      it "uses correct site" do
        expect(strategy.client.site).to eq("https://login.microsoftonline.com")
      end

      it "uses correct authorize_url" do
        expect(strategy.client.authorize_url).to eq("https://login.microsoftonline.com/common/oauth2/v2.0/authorize")
      end

      it "uses correct token_url" do
        expect(strategy.client.token_url).to eq("https://login.microsoftonline.com/common/oauth2/v2.0/token")
      end
    end

    context "with customized client options" do
      let(:options) do
        {
          client_options: {
            'site'          => 'https://example.com',
            'authorize_url' => 'https://example.com/authorize',
            'token_url'     => 'https://example.com/token',
          }
        }
      end

      it "uses customized site" do
        expect(strategy.client.site).to eq("https://example.com")
      end

      it "uses customized authorize_url" do
        expect(strategy.client.authorize_url).to eq("https://example.com/authorize")
      end

      it "uses customized token_url" do
        expect(strategy.client.token_url).to eq("https://example.com/token")
      end
    end
  end

  describe "#authorize_params" do
    let(:options) do
      { authorize_params: { foo: "bar", baz: "zip" } }
    end

    it "uses correct scope and allows to customize authorization parameters" do
      expect(strategy.authorize_params).to match(
        "scope" => "openid email profile https://outlook.office.com/contacts.read",
        "foo" => "bar",
        "baz" => "zip",
        "state" => /\A\h{48}\z/
      )
    end
  end

end
