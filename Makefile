NAME      := xllify
NS        := XLLIFY_START
BASE_URL  := https://localhost:3000
FUNCTIONS := $(wildcard functions/*.luau)

.PHONY: xll officejs dev clean

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

clean:
	rm -f builds/$(NAME).xll builds/$(NAME).zip builds/$(NAME)-dev.zip
