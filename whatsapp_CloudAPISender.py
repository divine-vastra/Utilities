import logging
import requests

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class WhatsAppCloudAPI:
    def __init__(self, token: str, phone_number_id: str, version: str = "v26.0"):
        self.base_url = f"https://graph.facebook.com/{version}/{phone_number_id}/messages"
        self.headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        }
        self.session = requests.Session()
        self.session.headers.update(self.headers)

    def build_image_header(self, url: str) -> dict:
        """Helper to structure an image header component."""
        return {
            "type": "header",
            "parameters": [
                {
                    "type": "image",
                    "image": {"link": url}
                }
            ]
        }

    def build_body_parameters(self, variables: list) -> dict:
        """Helper to structure body text variables (ordered list)."""
        return {
            "type": "body",
            "parameters": [{"type": "text", "text": str(v)} for v in variables]
        }

    def build_url_button(self, index: int, value: str) -> dict:
        """Helper to append a dynamic suffix/code to a URL button."""
        return {
            "type": "button",
            "sub_type": "url",
            "index": str(index),
            "parameters": [
                {
                    "type": "text",
                    "text": value
                }
            ]
        }

    def build_quick_reply_button(self, index: int, payload: str) -> dict:
        """Helper to attach a tracking payload to a Quick Reply button."""
        return {
            "type": "button",
            "sub_type": "quick_reply",
            "index": str(index),
            "parameters": [
                {
                    "type": "payload",
                    "payload": payload
                }
            ]
        }

    def send_template(self, to_phone: str, template_name: str, components: list, language_code: str = "en_US") -> dict:
        """
        Sends a rich media/button template message using assembled components.
        """
        payload = {
            "messaging_product": "whatsapp",
            "recipient_type": "individual",
            "to": to_phone,
            "type": "template",
            "template": {
                "name": template_name,
                "language": {"code": language_code},
                "components": components
            }
        }

        try:
            response = self.session.post(self.base_url, json=payload, timeout=10)
            response.raise_for_status()
            result = response.json()
            logger.info(f"Template message sent successfully to {to_phone}. ID: {result.get('messages', [{}])[0].get('id')}")
            return result
        except requests.exceptions.HTTPError as http_err:
            logger.error(f"HTTP error occurred: {http_err} - Response: {response.text}")
            raise
        except Exception as err:
            logger.error(f"An unexpected error occurred: {err}")
            raise

if __name__ == "__main__":
    # Initialize client with your permanent system user token
    TOKEN = "EAAP2eVZBJ4XEBSHzzxnwyNjPdY3U2MURgP1O6ejGVDDlckfIeEQy5PspppPkq0QXl38jaqsmM9Kt1s5BGVLZAxf8RYynoZCLFYqjCfIg4ZB1eoS2VCZB8LVEWboyg4EL4UeUbOdDsJ0cxrIZBFvZARm9doSYDjlEZCVVkJFfc4tDju8DpNRr97dVQRAAacAhBQZDZD" 
    PHONE_NUMBER_ID = "635535569653631"
    
    client = WhatsAppCloudAPI(token=TOKEN, phone_number_id=PHONE_NUMBER_ID)

    # 1. Assemble template parts step-by-step
    components = [
        #client.build_image_header("https://scontent.whatsapp.net/v/t61.29466-34/583059929_1028302110202052_7659873356647554648_n.png?ccb=1-7&_nc_sid=8b1bef&_nc_ohc=WocEQjIE8WUQ7kNvwG9ik2F&_nc_oc=AdoT0K4wZj2nfY6VMJnbHU8_dFTvNmuTBnzk3Q4sk7mYUIPrbVADIcFVIjHyyrC6quvfDVSbkbHF2qFq9tedHjpW&_nc_zt=3&_nc_ht=scontent.whatsapp.net&_nc_gid=lVXDTXcI2GzSihh2Kz9Irg&_nc_ss=782a8&oh=01_Q5Aa5QEUshSZ5EKIucYeRfsyxA6e6YvMz2SUEF4Wb2WO5E-M1w&oe=6AA53E41"),
        #client.build_body_parameters(["1091402940391"]),                 # Maps to {{1}} in template body
        client.build_url_button(index=0, value="7D137810152&cp_id=105"),     # Appends to URL button at index 0
        #client.build_quick_reply_button(index=1, payload="TRACK_ID_99213") # Hidden payload for index 1
    ]

    # 2. Dispatch payload
    try:
        response_data = client.send_template(
            to_phone="919971679624",
            template_name="out_for_delivery",
            components=components,
            language_code="en"
        )
        print("Success payload:", response_data)
    except Exception as e:
        print("Broadcast failed. Check logs for details.")