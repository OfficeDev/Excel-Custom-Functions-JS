/**
 * Add two numbers
 * @customfunction 
 * @param {number} first 
 * @param {number} second 
 * @returns {number} The sum of first and second.
 */
function add(first, second) {
  return first + second;
}

/**
 * Returns the current time once a second
 * @customfunction 
 * @param handler {CustomFunctions.StreamingHandler<string>} Custom function handler  
 */
function clock(handler) {
  const timer = setInterval(() => {
    const time = currentTime();
    handler.setResult(time);
  }, 1000);

  handler.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Returns the current time
 * @customfunction 
 * @returns {string} String containing the current time formatted for the current locale.
 */
function currentTime() {
  return new Date().toLocaleTimeString();
}

/**
 * Increments a number once a second.
 * @customfunction 
 * @param {number} incrementBy Amount to increment
 * @param {*} handler 
 */
function increment(incrementBy, handler) {
  let result = 0;
  const timer = setInterval(() => {
    result += incrementBy;
    handler.setResult(result);
  }, 1000);

  handler.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Writes a message to console.log().
 * @customfunction 
 * @param {string} message String to log
 */
function logMessage(message) {
  console.log(message);

  return message;
}

/**
 * Defines the implementation of the custom functions
 * for the function id defined in the metadata file (functions.json).
 */
CustomFunctions.associate("ADD", add);
CustomFunctions.associate("CLOCK", clock);
CustomFunctions.associate("INCREMENT", increment);
CustomFunctions.associate("LOG", logMessage);
