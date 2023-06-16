#define portLed 13

void setup() {
	Serial.begin(9600);
	Serial.println("Setup Finalizado...");
	pinMode(portLed, OUTPUT);

}

void loop() {
	digitalWrite(portLed, HIGH);
	delay(100);
	digitalWrite(portLed, LOW);
	delay(100);

}
