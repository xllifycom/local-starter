NAME      := xllify
NS        := XLLIFY_START
BASE_URL  := https://localhost:3000
FUNCTIONS := $(wildcard functions/*.luau)

.PHONY: xll officejs dev clean new-project

xll: builds/$(NAME).xll

officejs: builds/$(NAME).zip

dev: builds/$(NAME)-dev.zip

builds/$(NAME).xll: $(FUNCTIONS) | builds/
	xllify build xll --title "$(NAME)" --namespace $(NS) -o $@ $(FUNCTIONS)
	@echo "Done. Copy $@ to a Windows machine; see https://xllify.com/xll-need-to-know for next steps."

builds/$(NAME).zip: $(FUNCTIONS) | builds/
	xllify build officejs --title "$(NAME)" --namespace $(NS) -o $@ $(FUNCTIONS)

builds/$(NAME)-dev.zip: $(FUNCTIONS) | builds/
	xllify build officejs --title "$(NAME)" --namespace $(NS) --dev-url $(BASE_URL) -o $@ $(FUNCTIONS)

builds/:
	mkdir -p builds

new-project:
	@UUID=$$(uuidgen 2>/dev/null || python3 -c "import uuid; print(uuid.uuid4())" 2>/dev/null) && \
	  UUID=$$(echo "$$UUID" | tr '[:upper:]' '[:lower:]') && \
	  [ -n "$$UUID" ] || { echo "Could not generate a UUID. Visit https://www.uuidgenerator.net/ and paste the result into the app_id field in xllify.json"; exit 1; } && \
	  sed -i.bak "s/\"app_id\": \".*\"/\"app_id\": \"$$UUID\"/" xllify.json && \
	  rm -f xllify.json.bak && \
	  echo "app_id set to $$UUID"

clean:
	rm -f builds/$(NAME).xll builds/$(NAME).zip builds/$(NAME)-dev.zip
